import type { IGraphPermissionAuditService, ISharePointPermissionAuditService } from '../services';

export interface IPermissionAuditProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  graphPermissionAuditService: IGraphPermissionAuditService;
  groupExpansionBatchSize: number;
  groupedViewPreferenceKey: string;
  sharePointPermissionAuditService: ISharePointPermissionAuditService;
}
