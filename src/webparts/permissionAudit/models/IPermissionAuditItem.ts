export type PermissionAuditPrincipalType =
  | 'List'
  | 'ListItem'
  | 'LoadMore'
  | 'Site'
  | 'SharePointGroup'
  | 'SecurityGroup'
  | 'DistributionList'
  | 'Microsoft365Group'
  | 'User'
  | 'Unknown';

export type PermissionAuditSourceType =
  | 'DirectPermission'
  | 'DeferredGroupMembers'
  | 'SecurableObject'
  | 'SharePointGroupMember'
  | 'NestedGroupMember'
  | 'ExpandedGroupMember';

export interface IPermissionAuditLevel {
  id?: number;
  name: string;
  description?: string;
  hidden?: boolean;
  order?: number;
}

export interface IPermissionAuditGroupDetails {
  aadObjectId?: string;
  allowMembersEditMembership?: boolean;
  allowRequestToJoinLeave?: boolean;
  autoAcceptRequestToJoinLeave?: boolean;
  expansionError?: string;
  isHiddenInUi?: boolean;
  onlyAllowMembersViewMembership?: boolean;
  ownerTitle?: string;
  requestToJoinLeaveEmail?: string;
}

export interface IPermissionAuditPersonDetails {
  email?: string;
  jobTitle?: string;
  department?: string;
  sipAddress?: string;
  userPrincipalName?: string;
}

export interface IPermissionAuditObjectDetails {
  baseTemplate?: number;
  hidden?: boolean;
  itemId?: number;
  itemCount?: number;
  serverRelativeUrl?: string;
  url?: string;
}

export interface IPermissionAuditItemDetail {
  label: string;
  value: string;
}

export type PermissionAuditDeferredGroupMemberSource =
  | 'GraphGroup'
  | 'SharePointGroup';

export interface IPermissionAuditDeferredGroupMembersDetails {
  batchSize: number;
  graphGroupAadObjectId?: string;
  graphGroupLoginName?: string;
  graphNextLink?: string;
  graphRelationship?: string;
  parentGroupDepth: number;
  parentGroupKey: string;
  parentGroupPath: string[];
  parentGroupSourceType: PermissionAuditSourceType;
  sharePointGroupId?: number;
  sharePointNextOffset?: number;
  source: PermissionAuditDeferredGroupMemberSource;
  visitedSharePointGroupIds?: string[];
}

export interface IPermissionAuditItem {
  key: string;
  displayName: string;
  principalType: PermissionAuditPrincipalType;
  sourceType: PermissionAuditSourceType;
  permissionLevels: IPermissionAuditLevel[];
  depth: number;
  path: string[];
  parentKey?: string;
  principalId?: number;
  loginName?: string;
  description?: string;
  groupDetails?: IPermissionAuditGroupDetails;
  objectDetails?: IPermissionAuditObjectDetails;
  personDetails?: IPermissionAuditPersonDetails;
  details?: IPermissionAuditItemDetail[];
  deferredGroupMembersDetails?: IPermissionAuditDeferredGroupMembersDetails;
  children?: IPermissionAuditItem[];
}
