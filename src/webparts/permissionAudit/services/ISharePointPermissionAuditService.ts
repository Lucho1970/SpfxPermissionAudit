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
  groupExpansionBatchSize?: number;
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
  groupExpansionBatchSize?: number;
  expandGroups?: boolean;
  includeHiddenListsWithUniquePermissions?: boolean;
  includeListItemsWithUniquePermissions?: boolean;
  includeListsWithUniquePermissions?: boolean;
  onAuditUpdated?: PermissionAuditProgressHandler;
}

export interface ISharePointPermissionAuditService {
  expandPermissionAuditGroupsAsync(groups: IPermissionAuditItem[], groupExpansionBatchSize?: number): Promise<IPermissionAuditItem[]>;
  getCurrentSitePermissionGroupsAsync(): Promise<IPermissionAuditItem[]>;
  getCurrentSitePermissionGroupsWithMetadataAsync(expandGroups?: boolean, groupExpansionBatchSize?: number): Promise<ICurrentSitePermissionGroupsResult>;
  getListPermissionAuditItemsAsync(options: ISharePointPermissionAuditOptions): Promise<IPermissionAuditItem[]>;
  getPermissionAuditAsync(options: ISharePointPermissionAuditOptions): Promise<ICurrentSitePermissionGroupsResult>;
  loadDeferredGroupMembersAsync(loadMoreItem: IPermissionAuditItem): Promise<IPermissionAuditItem[]>;
  searchPrincipalAccessAsync(request: IPrincipalAccessSearchRequest): Promise<IPrincipalAccessSearchResult>;
}
