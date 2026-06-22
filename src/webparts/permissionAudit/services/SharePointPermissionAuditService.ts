import type { SPFI } from '@pnp/sp';
import type { IPermissionAuditItem, IPermissionAuditLevel } from '../models';
import type { IGraphPermissionAuditService } from './IGraphPermissionAuditService';
import type {
  IAssociatedSharePointGroupIds,
  ICurrentSitePermissionGroupsResult,
  ISharePointPermissionAuditOptions,
  ISharePointPermissionAuditService
} from './ISharePointPermissionAuditService';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/security';
import '@pnp/sp/security/item';
import '@pnp/sp/security/list';
import '@pnp/sp/site-groups';
import '@pnp/sp/site-users';

export interface ISharePointSiteGroupInfo {
  AllowMembersEditMembership: boolean;
  AllowRequestToJoinLeave: boolean;
  AutoAcceptRequestToJoinLeave: boolean;
  Description: string;
  Id: number;
  IsHiddenInUI: boolean;
  LoginName: string;
  OnlyAllowMembersViewMembership: boolean;
  OwnerTitle: string;
  PrincipalType: number;
  RequestToJoinLeaveEmailSetting?: string;
  Title: string;
}

interface IRoleDefinitionBindingInfo {
  Name: string;
}

interface IRoleAssignmentInfo {
  Member?: {
    Id: number;
    LoginName?: string;
    PrincipalType: number;
    Title?: string;
  };
  PrincipalId: number;
  RoleDefinitionBindings?: IRoleDefinitionBindingInfo[] | {
    results?: IRoleDefinitionBindingInfo[];
  };
}

interface ISharePointWebInfo {
  Description?: string;
  Id: string;
  ServerRelativeUrl?: string;
  Title: string;
  Url: string;
}

interface ISharePointListInfo {
  BaseTemplate: number;
  HasUniqueRoleAssignments: boolean;
  Description?: string;
  Hidden: boolean;
  Id: string;
  ItemCount: number;
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
  Title: string;
}

interface ISharePointListItemInfo {
  FileLeafRef?: string;
  FileRef?: string;
  Id: number;
  Title?: string;
}

interface ISharePointListItemPermissionInfo {
  HasUniqueRoleAssignments: boolean;
}

interface ISiteUserInfo {
  Email: string;
  Id: number;
  IsHiddenInUI: boolean;
  IsSiteAdmin: boolean;
  LoginName: string;
  PrincipalType: number;
  Title: string;
  UserId?: {
    NameId?: string;
    NameIdIssuer?: string;
  };
  UserPrincipalName?: string;
}

const sharePointGroupPrincipalType: number = 8;
const userPrincipalType: number = 1;
const distributionListPrincipalType: number = 2;
const securityGroupPrincipalType: number = 4;

const toDisplayText = (value: unknown): string => `${value ?? ''}`;

const hasPrincipalType = (principalType: number, expectedPrincipalType: number): boolean =>
  (principalType & expectedPrincipalType) === expectedPrincipalType;

const getRoleBindings = (assignment: IRoleAssignmentInfo): IRoleDefinitionBindingInfo[] => {
  const bindings = assignment.RoleDefinitionBindings;

  if (Array.isArray(bindings)) {
    return bindings;
  }

  return bindings?.results || [];
};

const toPermissionLevels = (assignment: IRoleAssignmentInfo | undefined): IPermissionAuditLevel[] => {
  if (!assignment) {
    return [];
  }

  return getRoleBindings(assignment)
    .filter((binding) => !!binding.Name)
    .map((binding) => ({
      name: binding.Name
    }));
};

const hasOnlyLimitedAccess = (permissionLevels: IPermissionAuditLevel[]): boolean =>
  permissionLevels.length === 1 && permissionLevels[0].name === 'Limited Access';

export class SharePointPermissionAuditService implements ISharePointPermissionAuditService {
  public constructor(
    private readonly _sp: SPFI,
    private readonly _graphPermissionAuditService: IGraphPermissionAuditService
  ) {
  }

  public async getCurrentSitePermissionGroupsAsync(): Promise<IPermissionAuditItem[]> {
    const result: ICurrentSitePermissionGroupsResult = await this.getCurrentSitePermissionGroupsWithMetadataAsync();

    return result.groups;
  }

  public async getCurrentSitePermissionGroupsWithMetadataAsync(expandGroups?: boolean): Promise<ICurrentSitePermissionGroupsResult> {
    const result: ICurrentSitePermissionGroupsResult = await this.getPermissionAuditAsync({ expandGroups });
    const siteNode: IPermissionAuditItem | undefined = result.groups[0];

    return {
      groups: siteNode?.children || [],
      associatedGroupIds: result.associatedGroupIds
    };
  }

  public async getPermissionAuditAsync(options: ISharePointPermissionAuditOptions): Promise<ICurrentSitePermissionGroupsResult> {
    const [
      webInfo,
      groups,
      roleAssignments,
      associatedOwnerGroup,
      associatedMemberGroup,
      associatedVisitorGroup
    ] = await Promise.all([
      this._sp.web.select('Description', 'Id', 'ServerRelativeUrl', 'Title', 'Url')(),
      this._sp.web.siteGroups(),
      this._sp.web.roleAssignments.expand('Member', 'RoleDefinitionBindings')(),
      this._sp.web.associatedOwnerGroup().catch(() => undefined),
      this._sp.web.associatedMemberGroup().catch(() => undefined),
      this._sp.web.associatedVisitorGroup().catch(() => undefined)
    ]) as [
      ISharePointWebInfo,
      ISharePointSiteGroupInfo[],
      IRoleAssignmentInfo[],
      ISharePointSiteGroupInfo | undefined,
      ISharePointSiteGroupInfo | undefined,
      ISharePointSiteGroupInfo | undefined
    ];

    const roleAssignmentByPrincipalId: Map<number, IRoleAssignmentInfo> = roleAssignments
      .filter((assignment) => !!assignment.Member && hasPrincipalType(assignment.Member.PrincipalType, sharePointGroupPrincipalType))
      .reduce((assignmentMap: Map<number, IRoleAssignmentInfo>, assignment: IRoleAssignmentInfo) => {
        assignmentMap.set(assignment.PrincipalId, assignment);

        return assignmentMap;
      }, new Map<number, IRoleAssignmentInfo>());

    const associatedGroupIds: IAssociatedSharePointGroupIds = {
      ownerGroupId: associatedOwnerGroup?.Id,
      memberGroupId: associatedMemberGroup?.Id,
      visitorGroupId: associatedVisitorGroup?.Id
    };

    let permissionGroups: IPermissionAuditItem[] = groups.map((group) => this._toPermissionAuditItem(
        group,
        toPermissionLevels(roleAssignmentByPrincipalId.get(group.Id))
      ));
    let directSitePermissions: IPermissionAuditItem[] = await Promise.all(roleAssignments
      .filter((assignment) => {
        if (!assignment.Member || hasPrincipalType(assignment.Member.PrincipalType, sharePointGroupPrincipalType)) {
          return false;
        }

        return !hasOnlyLimitedAccess(toPermissionLevels(assignment));
      })
      .map((assignment) =>
        this._toPermissionAuditItemFromRoleAssignmentAsync(
          assignment,
          `site-${webInfo.Id}`,
          toDisplayText(webInfo.Title),
          [toDisplayText(webInfo.Title)],
          1,
          options.expandGroups
        )
      ));

    if (options.expandGroups) {
      [permissionGroups, directSitePermissions] = await Promise.all([
        Promise.all(permissionGroups.map((permissionGroup) =>
          this._withExpandedPermissionItemGroupAsync(permissionGroup, new Set<string>([permissionGroup.key]))
        )),
        Promise.resolve(directSitePermissions)
      ]);
    }

    const siteNode: IPermissionAuditItem = this._toSiteAuditItem(webInfo, [...directSitePermissions, ...permissionGroups]);
    const result: ICurrentSitePermissionGroupsResult = {
      groups: [siteNode],
      associatedGroupIds
    };

    if (options.includeListsWithUniquePermissions || options.includeListItemsWithUniquePermissions) {
      const lists: ISharePointListInfo[] = await this._getListsForPermissionAuditAsync(options);
      const displayedListNodeById: Map<string, IPermissionAuditItem> = new Map<string, IPermissionAuditItem>();
      const listNodes: IPermissionAuditItem[] = options.includeListsWithUniquePermissions
        ? await Promise.all(lists
          .filter((list) => list.HasUniqueRoleAssignments)
          .map((list) => this._toListAuditItemAsync(list, options, true)))
        : [];

      listNodes.forEach((listNode) => {
        displayedListNodeById.set(listNode.key.replace(/^list-/, ''), listNode);
      });

      siteNode.children = [...(siteNode.children || []), ...listNodes];
      options.onAuditUpdated?.(result);

      if (options.includeListItemsWithUniquePermissions) {
        await Promise.all(lists.map(async (list) => {
          let listNode: IPermissionAuditItem | undefined = displayedListNodeById.get(list.Id);

          if (!listNode) {
            listNode = this._toListContainerAuditItem(list);
          }

          const listItems: IPermissionAuditItem[] = await this._getListItemsWithUniquePermissionsAsync(listNode, options);

          if (listItems.length > 0) {
            listNode.children = [...(listNode.children || []), ...listItems];

            if (!displayedListNodeById.has(list.Id)) {
              displayedListNodeById.set(list.Id, listNode);
              siteNode.children = [...(siteNode.children || []), listNode];
            }

            options.onAuditUpdated?.(result);
          }
        }));
      }
    }

    return result;
  }

  private async _getListsForPermissionAuditAsync(options: ISharePointPermissionAuditOptions): Promise<ISharePointListInfo[]> {
    const listFilter: string[] = [];

    if (options.includeListsWithUniquePermissions && !options.includeListItemsWithUniquePermissions) {
      listFilter.push('HasUniqueRoleAssignments eq true');
    }

    if (!options.includeHiddenListsWithUniquePermissions) {
      listFilter.push('Hidden eq false');
    }

    const listQuery = this._sp.web.lists
      .select('BaseTemplate', 'Description', 'Hidden', 'Id', 'ItemCount', 'RootFolder/ServerRelativeUrl', 'Title', 'HasUniqueRoleAssignments')
      .expand('RootFolder')
      .top(5000);

    if (listFilter.length === 0) {
      return listQuery() as Promise<ISharePointListInfo[]>;
    }

    return listQuery.filter(listFilter.join(' and '))() as Promise<ISharePointListInfo[]>;
  }

  private async _toListAuditItemAsync(
    list: ISharePointListInfo,
    options: ISharePointPermissionAuditOptions,
    includePermissionChildren: boolean
  ): Promise<IPermissionAuditItem> {
    const listNode: IPermissionAuditItem = this._toListContainerAuditItem(list);

    if (!includePermissionChildren) {
      return listNode;
    }

    const roleAssignments: IRoleAssignmentInfo[] = await this._sp.web.lists
      .getById(list.Id)
      .roleAssignments
      .expand('Member', 'RoleDefinitionBindings')();
    let permissionChildren: IPermissionAuditItem[] = await Promise.all(roleAssignments.map((assignment) =>
      this._toPermissionAuditItemFromRoleAssignmentAsync(
        assignment,
        listNode.key,
        listNode.displayName,
        listNode.path,
        listNode.depth + 1,
        options.expandGroups
      )
    ));

    permissionChildren = permissionChildren.filter((permissionChild) => permissionChild.permissionLevels.length > 0);

    return {
      ...listNode,
      children: permissionChildren
    };
  }

  private _toListContainerAuditItem(list: ISharePointListInfo): IPermissionAuditItem {
    return {
      key: `list-${list.Id}`,
      displayName: toDisplayText(list.Title),
      principalType: 'List',
      sourceType: 'SecurableObject',
      permissionLevels: [],
      depth: 1,
      path: [toDisplayText(list.Title)],
      description: toDisplayText(list.Description),
      objectDetails: {
        baseTemplate: list.BaseTemplate,
        hidden: list.Hidden,
        itemCount: list.ItemCount,
        serverRelativeUrl: toDisplayText(list.RootFolder?.ServerRelativeUrl)
      },
      children: []
    };
  }

  private async _getListItemsWithUniquePermissionsAsync(
    listNode: IPermissionAuditItem,
    options: ISharePointPermissionAuditOptions
  ): Promise<IPermissionAuditItem[]> {
    const listItems: ISharePointListItemInfo[] = await this._sp.web.lists
      .getById(listNode.key.replace(/^list-/, ''))
      .items
      .select('FileLeafRef', 'FileRef', 'Id', 'Title')
      .top(5000)() as ISharePointListItemInfo[];

    const auditedListItems: Array<IPermissionAuditItem | undefined> = await Promise.all(listItems.map(async (listItem) => {
      const permissionInfo: ISharePointListItemPermissionInfo = await this._sp.web.lists
        .getById(listNode.key.replace(/^list-/, ''))
        .items
        .getById(listItem.Id)
        .select('HasUniqueRoleAssignments')() as ISharePointListItemPermissionInfo;

      if (!permissionInfo.HasUniqueRoleAssignments) {
        return undefined;
      }

      return this._toListItemAuditItemAsync(listNode, listItem, options);
    }));

    return auditedListItems.filter((listItem): listItem is IPermissionAuditItem => !!listItem);
  }

  private async _toListItemAuditItemAsync(
    listNode: IPermissionAuditItem,
    listItem: ISharePointListItemInfo,
    options: ISharePointPermissionAuditOptions
  ): Promise<IPermissionAuditItem> {
    const listId: string = listNode.key.replace(/^list-/, '');
    const itemDisplayName: string = toDisplayText(listItem.FileLeafRef || listItem.Title || `Item ${listItem.Id}`);
    const itemKey: string = `${listNode.key}|item-${listItem.Id}`;
    const itemPath: string[] = [...listNode.path, itemDisplayName];
    const roleAssignments: IRoleAssignmentInfo[] = await this._sp.web.lists
      .getById(listId)
      .items
      .getById(listItem.Id)
      .roleAssignments
      .expand('Member', 'RoleDefinitionBindings')();
    let permissionChildren: IPermissionAuditItem[] = await Promise.all(roleAssignments.map((assignment) =>
      this._toPermissionAuditItemFromRoleAssignmentAsync(
        assignment,
        itemKey,
        itemDisplayName,
        itemPath,
        listNode.depth + 2,
        options.expandGroups
      )
    ));

    permissionChildren = permissionChildren.filter((permissionChild) =>
      permissionChild.permissionLevels.length > 0 && !hasOnlyLimitedAccess(permissionChild.permissionLevels)
    );

    return {
      key: itemKey,
      displayName: itemDisplayName,
      principalType: 'ListItem',
      sourceType: 'SecurableObject',
      permissionLevels: [],
      depth: listNode.depth + 1,
      path: itemPath,
      parentKey: listNode.key,
      objectDetails: {
        itemId: listItem.Id,
        serverRelativeUrl: toDisplayText(listItem.FileRef)
      },
      children: permissionChildren
    };
  }

  private async _toPermissionAuditItemFromRoleAssignmentAsync(
    assignment: IRoleAssignmentInfo,
    parentKey: string,
    parentDisplayName: string,
    parentPath: string[],
    depth: number,
    expandGroups?: boolean
  ): Promise<IPermissionAuditItem> {
    const member = assignment.Member;
    const permissionLevels: IPermissionAuditLevel[] = toPermissionLevels(assignment);
    const displayName: string = toDisplayText(member?.Title || member?.LoginName || assignment.PrincipalId);
    const baseItem: IPermissionAuditItem = {
      key: `${parentKey}|principal-${assignment.PrincipalId}`,
      displayName,
      principalType: member ? this._getPrincipalType(member) : 'Unknown',
      sourceType: 'DirectPermission',
      permissionLevels,
      depth,
      path: [...parentPath, displayName],
      parentKey,
      principalId: assignment.PrincipalId,
      loginName: toDisplayText(member?.LoginName),
      children: []
    };

    if (!expandGroups || !member) {
      return baseItem;
    }

    if (hasPrincipalType(member.PrincipalType, sharePointGroupPrincipalType)) {
      return this._withExpandedPermissionItemGroupAsync(baseItem, new Set<string>([parentKey, assignment.PrincipalId.toString()]));
    }

    if (hasPrincipalType(member.PrincipalType, securityGroupPrincipalType) ||
      hasPrincipalType(member.PrincipalType, distributionListPrincipalType)) {
      const children: IPermissionAuditItem[] = await this._graphPermissionAuditService.expandGroupMembersAsync({
        groupAadObjectId: this._extractAadObjectIdFromLoginName(member.LoginName),
        groupDisplayName: displayName,
        groupLoginName: member.LoginName,
        inheritedPermissionLevels: permissionLevels,
        parentKey: baseItem.key,
        parentPath: baseItem.path,
        depth: baseItem.depth + 1
      });

      return {
        ...baseItem,
        children
      };
    }

    return baseItem;
  }

  private async _withExpandedPermissionItemGroupAsync(
    permissionGroup: IPermissionAuditItem,
    visitedGroupKeys: Set<string>
  ): Promise<IPermissionAuditItem> {
    if (!permissionGroup.principalId) {
      return permissionGroup;
    }

    const members: ISiteUserInfo[] = await this._sp.web.siteGroups
      .getById(permissionGroup.principalId)
      .users
      .select('Email', 'Id', 'IsHiddenInUI', 'IsSiteAdmin', 'LoginName', 'PrincipalType', 'Title', 'UserId', 'UserPrincipalName')
      .top(5000)();
    const childItems: IPermissionAuditItem[] = await Promise.all(members.map((member) =>
      this._toPermissionAuditMemberItemAsync(member, permissionGroup, visitedGroupKeys)
    ));

    return {
      ...permissionGroup,
      children: childItems
    };
  }

  private async _toPermissionAuditMemberItemAsync(
    member: ISiteUserInfo,
    parentGroup: IPermissionAuditItem,
    visitedGroupKeys: Set<string>
  ): Promise<IPermissionAuditItem> {
    const memberKey: string = `${parentGroup.key}|${member.PrincipalType}|${member.Id}`;
    const displayName: string = toDisplayText(member.Title || member.UserPrincipalName || member.Email || member.LoginName);
    const inheritedPermissionLevels: IPermissionAuditLevel[] = parentGroup.permissionLevels;
    const baseItem: IPermissionAuditItem = {
      key: memberKey,
      displayName,
      principalType: this._getPrincipalType(member),
      sourceType: parentGroup.sourceType === 'DirectPermission' ? 'SharePointGroupMember' : 'NestedGroupMember',
      permissionLevels: inheritedPermissionLevels,
      depth: parentGroup.depth + 1,
      path: [...parentGroup.path, displayName],
      parentKey: parentGroup.key,
      principalId: member.Id,
      loginName: toDisplayText(member.LoginName),
      personDetails: hasPrincipalType(member.PrincipalType, userPrincipalType) ? {
        email: toDisplayText(member.Email),
        userPrincipalName: toDisplayText(member.UserPrincipalName)
      } : undefined,
      groupDetails: !hasPrincipalType(member.PrincipalType, userPrincipalType) ? {
        aadObjectId: this._extractAadObjectId(member),
        isHiddenInUi: member.IsHiddenInUI
      } : undefined,
      children: []
    };

    if (hasPrincipalType(member.PrincipalType, sharePointGroupPrincipalType) && !visitedGroupKeys.has(member.Id.toString())) {
      const nextVisitedGroupKeys: Set<string> = new Set(visitedGroupKeys);
      nextVisitedGroupKeys.add(member.Id.toString());

      return this._withExpandedPermissionItemGroupAsync(baseItem, nextVisitedGroupKeys);
    }

    if (hasPrincipalType(member.PrincipalType, securityGroupPrincipalType) ||
      hasPrincipalType(member.PrincipalType, distributionListPrincipalType)) {
      const children: IPermissionAuditItem[] = await this._graphPermissionAuditService.expandGroupMembersAsync({
        groupAadObjectId: baseItem.groupDetails?.aadObjectId,
        groupDisplayName: displayName,
        groupLoginName: member.LoginName,
        inheritedPermissionLevels,
        parentKey: memberKey,
        parentPath: baseItem.path,
        depth: baseItem.depth + 1
      });

      return {
        ...baseItem,
        children
      };
    }

    return baseItem;
  }

  private _toPermissionAuditItem(
    group: ISharePointSiteGroupInfo,
    permissionLevels: IPermissionAuditLevel[]
  ): IPermissionAuditItem {
    return {
      key: group.Id.toString(),
      displayName: toDisplayText(group.Title),
      principalType: 'SharePointGroup',
      sourceType: 'DirectPermission',
      permissionLevels,
      depth: 0,
      path: [toDisplayText(group.Title)],
      principalId: group.Id,
      loginName: toDisplayText(group.LoginName),
      description: toDisplayText(group.Description),
      groupDetails: {
        allowMembersEditMembership: group.AllowMembersEditMembership,
        allowRequestToJoinLeave: group.AllowRequestToJoinLeave,
        autoAcceptRequestToJoinLeave: group.AutoAcceptRequestToJoinLeave,
        isHiddenInUi: group.IsHiddenInUI,
        onlyAllowMembersViewMembership: group.OnlyAllowMembersViewMembership,
        ownerTitle: toDisplayText(group.OwnerTitle),
        requestToJoinLeaveEmail: toDisplayText(group.RequestToJoinLeaveEmailSetting)
      },
      children: []
    };
  }

  private _toSiteAuditItem(webInfo: ISharePointWebInfo, children: IPermissionAuditItem[]): IPermissionAuditItem {
    const siteTitle: string = toDisplayText(webInfo.Title);

    return {
      key: `site-${webInfo.Id}`,
      displayName: siteTitle,
      principalType: 'Site',
      sourceType: 'SecurableObject',
      permissionLevels: [],
      depth: 0,
      path: [siteTitle],
      description: toDisplayText(webInfo.Description),
      objectDetails: {
        serverRelativeUrl: toDisplayText(webInfo.ServerRelativeUrl),
        url: toDisplayText(webInfo.Url)
      },
      children: children.map((child) => this._withParentContext(child, `site-${webInfo.Id}`, [siteTitle], 1))
    };
  }

  private _withParentContext(
    permissionItem: IPermissionAuditItem,
    parentKey: string,
    parentPath: string[],
    depth: number
  ): IPermissionAuditItem {
    const path: string[] = [...parentPath, permissionItem.displayName];

    return {
      ...permissionItem,
      depth,
      parentKey,
      path,
      children: permissionItem.children?.map((child) => this._withParentContext(child, permissionItem.key, path, depth + 1))
    };
  }

  private _getPrincipalType(member: { PrincipalType: number }): IPermissionAuditItem['principalType'] {
    switch (member.PrincipalType) {
      case sharePointGroupPrincipalType:
        return 'SharePointGroup';
      default:
        if (hasPrincipalType(member.PrincipalType, userPrincipalType)) {
          return 'User';
        }

        if (hasPrincipalType(member.PrincipalType, distributionListPrincipalType)) {
          return 'DistributionList';
        }

        if (hasPrincipalType(member.PrincipalType, securityGroupPrincipalType)) {
          return 'SecurityGroup';
        }

        return 'Unknown';
    }
  }

  private _extractAadObjectId(member: ISiteUserInfo): string | undefined {
    const objectIdCandidates: (string | undefined)[] = [
      member.UserId?.NameId,
      member.LoginName?.split('|').pop()
    ];

    return objectIdCandidates
      .map((objectIdCandidate) =>
        objectIdCandidate?.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/)?.[0]
      )
      .find((objectIdCandidate) => !!objectIdCandidate);
  }

  private _extractAadObjectIdFromLoginName(loginName: string | undefined): string | undefined {
    return loginName?.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/)?.[0];
  }
}
