import type { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import type { IPermissionAuditItem } from '../models';
import type { IGraphGroupExpansionRequest, IGraphPermissionAuditService } from './IGraphPermissionAuditService';

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

export class GraphPermissionAuditService implements IGraphPermissionAuditService {
  public constructor(private readonly _msGraphClientFactory?: MSGraphClientFactory) {
  }

  public async expandGroupMembersAsync(request: IGraphGroupExpansionRequest): Promise<IPermissionAuditItem[]> {
    const aadObjectId: string | undefined = request.groupAadObjectId || extractAadObjectId(request.groupLoginName);

    if (!aadObjectId || !this._msGraphClientFactory) {
      return [this._toExpansionErrorItem(
        request,
        `Unable to expand this directory group because its Entra object id was not available from the SharePoint principal. Login name: ${toDisplayText(request.groupLoginName)}.`
      )];
    }

    let members: IGraphDirectoryObject[] = [];

    try {
      const client: MSGraphClientV3 = await this._msGraphClientFactory.getClient('3');
      members = await this._getGroupPrincipalsAsync(client, aadObjectId, getGroupClaimRole(request.groupLoginName));
    } catch (error) {
      return [this._toExpansionErrorItem(
        request,
        `Unable to expand this directory group through Microsoft Graph. ${error instanceof Error ? error.message : ''}`
      )];
    }

    return members.map((member) => ({
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
    }));
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

  private async _getGroupPrincipalsAsync(
    client: MSGraphClientV3,
    aadObjectId: string,
    claimRole: GraphGroupClaimRole
  ): Promise<IGraphDirectoryObject[]> {
    const members: IGraphDirectoryObject[] = [];
    const relationship: string = claimRole === 'Owners' ? 'owners' : 'transitiveMembers';
    let path: string | undefined =
      `/groups/${aadObjectId}/${relationship}?$select=id,displayName,mail,userPrincipalName,jobTitle,department&$top=999`;

    while (path) {
      const response: IGraphCollectionResponse<IGraphDirectoryObject> = await client.api(path).version('v1.0').get();
      members.push(...response.value);
      path = response['@odata.nextLink'];
    }

    return members;
  }
}
