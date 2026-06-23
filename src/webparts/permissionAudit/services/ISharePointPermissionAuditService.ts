import type { IPermissionAuditItem } from '../models';
import type { IDirectoryPersonInfo } from './IGraphPermissionAuditService';

export interface IAssociatedSharePointGroupIds {
  ownerGroupId?: number;
  memberGroupId?: number;
  visitorGroupId?: number;
}

export interface ICurrentSitePermissionGroupsResult {
  groups: IPermissionAuditItem[];
  associatedGroupIds: IAssociatedSharePointGroupIds;
}

export interface IPrincipalAccessMatch {
  matchItem: IPermissionAuditItem;
  scopeItem: IPermissionAuditItem;
}

export interface IPrincipalAccessSearchRequest {
  principal: IDirectoryPersonInfo;
  includeHiddenListsWithUniquePermissions?: boolean;
}

export interface IPrincipalAccessSearchResult {
  principal: IDirectoryPersonInfo;
  auditResult: ICurrentSitePermissionGroupsResult;
  matches: IPrincipalAccessMatch[];
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
  searchPrincipalAccessAsync(request: IPrincipalAccessSearchRequest): Promise<IPrincipalAccessSearchResult>;
}
