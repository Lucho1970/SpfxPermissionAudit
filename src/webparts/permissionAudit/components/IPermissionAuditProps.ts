import type { SPFI } from '@pnp/sp';

export interface IPermissionAuditProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
}
