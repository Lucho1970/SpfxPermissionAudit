import type { IPermissionAuditItem, IPermissionAuditLevel } from '../models';

export interface IGraphGroupExpansionRequest {
  groupDisplayName: string;
  groupAadObjectId?: string;
  groupLoginName?: string;
  inheritedPermissionLevels: IPermissionAuditLevel[];
  parentKey: string;
  parentPath: string[];
  depth: number;
}

export interface IGraphPermissionAuditService {
  expandGroupMembersAsync(request: IGraphGroupExpansionRequest): Promise<IPermissionAuditItem[]>;
}
