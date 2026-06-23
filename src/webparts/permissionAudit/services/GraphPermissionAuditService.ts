import type { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import type { IPermissionAuditItem } from '../models';
import type { IDirectoryPersonInfo, IGraphGroupExpansionRequest, IGraphPermissionAuditService } from './IGraphPermissionAuditService';

interface IGraphDirectoryObject {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  '@odata.type'?: string;
}

interface IGraphCollectionResponse<TItem> {
  value: TItem[];
  '@odata.nextLink'?: string;
}

type GraphGroupClaimRole = 'Members' | 'Owners';

const defaultGroupExpansionBatchSize: number = 100;

const toDisplayText = (value: unknown): string => `${value ?? ''}`;

const getFirstGuid = (value: string | undefined): string | undefined =>
  value?.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/)?.[0];

const extractAadObjectId = (loginName: string | undefined): string | undefined => {
  return getFirstGuid(loginName);
};

const isGraphGroup = (directoryObject: IGraphDirectoryObject): boolean =>
  directoryObject['@odata.type'] === '#microsoft.graph.group';

const getGroupClaimSuffix = (loginName: string | undefined): string | undefined => {
  const claimValue: string | undefined = loginName?.split('|').pop()?.toLowerCase();

  return claimValue?.match(/_([a-z])$/)?.[1];
};

const getGroupClaimRole = (loginName: string | undefined): GraphGroupClaimRole => {
  switch (getGroupClaimSuffix(loginName)) {
    case 'o':
      return 'Owners';
    case 'm':
    default:
      return 'Members';
  }
};

const toGraphApiPath = (pathOrUrl: string): string =>
  pathOrUrl.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, '');

export class GraphPermissionAuditService implements IGraphPermissionAuditService {
  public constructor(private readonly _msGraphClientFactory?: MSGraphClientFactory) {
  }

  public async searchPeopleAsync(searchText: string): Promise<IDirectoryPersonInfo[]> {
    const normalizedSearchText: string = searchText.trim();

    if (normalizedSearchText.length < 3 || !this._msGraphClientFactory) {
      return [];
    }

    const escapedSearchText: string = normalizedSearchText.replace(/'/g, "''");
    const filter: string = [
      `startswith(displayName,'${escapedSearchText}')`,
      `startswith(mail,'${escapedSearchText}')`,
      `startswith(userPrincipalName,'${escapedSearchText}')`
    ].join(' or ');
    const path: string =
      `/users?$select=id,displayName,mail,userPrincipalName,jobTitle,department&$top=8&$filter=${filter}`;
    const client: MSGraphClientV3 = await this._msGraphClientFactory.getClient('3');
    const response: IGraphCollectionResponse<IGraphDirectoryObject> = await client.api(path).version('v1.0').get();

    return response.value
      .filter((person) => !!person.id)
      .map((person) => ({
        id: person.id,
        displayName: toDisplayText(person.displayName || person.userPrincipalName || person.mail || person.id),
        email: toDisplayText(person.mail),
        userPrincipalName: toDisplayText(person.userPrincipalName),
        jobTitle: toDisplayText(person.jobTitle),
        department: toDisplayText(person.department)
      }));
  }

  public async expandGroupMembersAsync(request: IGraphGroupExpansionRequest): Promise<IPermissionAuditItem[]> {
    const aadObjectId: string | undefined = request.groupAadObjectId || extractAadObjectId(request.groupLoginName);

    if (!aadObjectId || !this._msGraphClientFactory) {
      return [this._toExpansionErrorItem(
        request,
        `Unable to expand this directory group because its Entra object id was not available from the SharePoint principal. Login name: ${toDisplayText(request.groupLoginName)}.`
      )];
    }

    let response: IGraphCollectionResponse<IGraphDirectoryObject>;

    try {
      const client: MSGraphClientV3 = await this._msGraphClientFactory.getClient('3');
      response = await this._getGroupPrincipalsPageAsync(
        client,
        aadObjectId,
        getGroupClaimRole(request.groupLoginName),
        request.batchSize || defaultGroupExpansionBatchSize
      );
    } catch (error) {
      return [this._toExpansionErrorItem(
        request,
        `Unable to expand this directory group through Microsoft Graph. ${error instanceof Error ? error.message : ''}`
      )];
    }

    return [
      ...response.value.map((member) => this._toDirectoryObjectAuditItem(member, request)),
      ...this._toDeferredGroupMembersItems(request, response['@odata.nextLink'], aadObjectId)
    ];
  }

  public async loadDeferredGroupMembersAsync(loadMoreItem: IPermissionAuditItem): Promise<IPermissionAuditItem[]> {
    const deferredDetails = loadMoreItem.deferredGroupMembersDetails;

    if (!deferredDetails?.graphNextLink || !this._msGraphClientFactory) {
      return [];
    }

    const request: IGraphGroupExpansionRequest = {
      batchSize: deferredDetails.batchSize,
      groupAadObjectId: deferredDetails.graphGroupAadObjectId,
      groupDisplayName: loadMoreItem.displayName,
      groupLoginName: deferredDetails.graphGroupLoginName,
      inheritedPermissionLevels: loadMoreItem.permissionLevels,
      parentKey: deferredDetails.parentGroupKey,
      parentPath: deferredDetails.parentGroupPath,
      depth: deferredDetails.parentGroupDepth + 1
    };
    const client: MSGraphClientV3 = await this._msGraphClientFactory.getClient('3');
    const members: IGraphDirectoryObject[] = [];
    let nextLink: string | undefined = deferredDetails.graphNextLink;

    while (nextLink) {
      const response: IGraphCollectionResponse<IGraphDirectoryObject> =
        await client.api(toGraphApiPath(nextLink)).version('v1.0').get();

      members.push(...response.value);
      nextLink = response['@odata.nextLink'];
    }

    return members.map((member) => this._toDirectoryObjectAuditItem(member, request));
  }

  private _toDirectoryObjectAuditItem(
    member: IGraphDirectoryObject,
    request: IGraphGroupExpansionRequest
  ): IPermissionAuditItem {
    return {
      key: `${request.parentKey}|graph|${member.id}`,
      displayName: toDisplayText(member.displayName || member.userPrincipalName || member.mail || member.id),
      principalType: isGraphGroup(member) ? 'SecurityGroup' : 'User',
      sourceType: 'ExpandedGroupMember',
      permissionLevels: request.inheritedPermissionLevels,
      depth: request.depth,
      path: [...request.parentPath, toDisplayText(member.displayName || member.userPrincipalName || member.mail || member.id)],
      parentKey: request.parentKey,
      loginName: toDisplayText(member.userPrincipalName || member.mail),
      groupDetails: isGraphGroup(member) ? {
        aadObjectId: member.id
      } : undefined,
      personDetails: isGraphGroup(member) ? undefined : {
        email: toDisplayText(member.mail),
        userPrincipalName: toDisplayText(member.userPrincipalName),
        jobTitle: toDisplayText(member.jobTitle),
        department: toDisplayText(member.department)
      },
      children: []
    };
  }

  private _toDeferredGroupMembersItems(
    request: IGraphGroupExpansionRequest,
    nextLink: string | undefined,
    aadObjectId: string | undefined
  ): IPermissionAuditItem[] {
    if (!nextLink) {
      return [];
    }

    const batchSize: number = request.batchSize || defaultGroupExpansionBatchSize;

    return [{
      key: `${request.parentKey}|graph-load-more|${encodeURIComponent(nextLink)}`,
      displayName: 'Load remaining group members',
      principalType: 'LoadMore',
      sourceType: 'DeferredGroupMembers',
      permissionLevels: request.inheritedPermissionLevels,
      depth: request.depth,
      path: [...request.parentPath, 'Load remaining group members'],
      parentKey: request.parentKey,
      deferredGroupMembersDetails: {
        batchSize,
        graphGroupAadObjectId: aadObjectId,
        graphGroupLoginName: request.groupLoginName,
        graphNextLink: nextLink,
        parentGroupDepth: request.depth - 1,
        parentGroupKey: request.parentKey,
        parentGroupPath: request.parentPath,
        parentGroupSourceType: 'ExpandedGroupMember',
        source: 'GraphGroup'
      },
      children: []
    }];
  }

  private _toExpansionErrorItem(request: IGraphGroupExpansionRequest, message: string): IPermissionAuditItem {
    return {
      key: `${request.parentKey}|graph-expansion-error`,
      displayName: request.groupDisplayName,
      principalType: 'Unknown',
      sourceType: 'ExpandedGroupMember',
      permissionLevels: request.inheritedPermissionLevels,
      depth: request.depth,
      path: [...request.parentPath, message],
      parentKey: request.parentKey,
      groupDetails: {
        expansionError: message
      },
      children: []
    };
  }

  private async _getGroupPrincipalsPageAsync(
    client: MSGraphClientV3,
    aadObjectId: string,
    claimRole: GraphGroupClaimRole,
    batchSize: number
  ): Promise<IGraphCollectionResponse<IGraphDirectoryObject>> {
    const relationship: string = claimRole === 'Owners' ? 'owners' : 'transitiveMembers';
    const path: string =
      `/groups/${aadObjectId}/${relationship}?$select=id,displayName,mail,userPrincipalName,jobTitle,department&$top=${batchSize}`;

    return client.api(path).version('v1.0').get();
  }
}
