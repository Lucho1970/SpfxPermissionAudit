import type { IPermissionAuditItem } from '../models';

export interface IAssociatedSharePointGroupIds {
  ownerGroupId?: number;
  memberGroupId?: number;
  visitorGroupId?: number;
}

export interface ICurrentSitePermissionGroupsResult {
  groups: IPermissionAuditItem[];
  associatedGroupIds: IAssociatedSharePointGroupIds;
}

export type PermissionAuditProgressHandler = (result: ICurrentSitePermissionGroupsResult) => void;

export interface ISharePointPermissionAuditOptions {
  expandGroups?: boolean;
  includeHiddenListsWithUniquePermissions?: boolean;
  includeListItemsWithUniquePermissions?: boolean;
  includeListsWithUniquePermissions?: boolean;
  onAuditUpdated?: PermissionAuditProgressHandler;
}

export interface ISharePointPermissionAuditService {
  getCurrentSitePermissionGroupsAsync(): Promise<IPermissionAuditItem[]>;
  getCurrentSitePermissionGroupsWithMetadataAsync(expandGroups?: boolean): Promise<ICurrentSitePermissionGroupsResult>;
  getPermissionAuditAsync(options: ISharePointPermissionAuditOptions): Promise<ICurrentSitePermissionGroupsResult>;
}
