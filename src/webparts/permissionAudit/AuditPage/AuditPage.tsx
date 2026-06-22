import * as React from 'react';
import { ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, IDetailsHeaderProps, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { IGroup, IGroupHeaderProps } from '@fluentui/react/lib/GroupedList';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Selection } from '@fluentui/react/lib/Selection';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { CheckboxList, ICheckboxListOption } from '../../../common/controls/CheckboxList';
import { HtmlContentDisplay } from '../../../common/controls/HtmlContentDisplay';
import type { IAuditPageProps } from './IAuditPageProps';
import type { IPermissionAuditItem, IPermissionAuditLevel } from '../models';
import type { IAssociatedSharePointGroupIds, ICurrentSitePermissionGroupsResult, ISharePointPermissionAuditOptions } from '../services';
import styles from './AuditPage.module.scss';
import * as strings from 'PermissionAuditWebPartStrings';

interface IAuditGridItem {
  key: string;
  permissionItem: IPermissionAuditItem;
  groupName: string;
  principalType: string;
  permissionLevels: string;
  groupPath: string;
  owner: string;
  hiddenInUi: string;
  membersCanEdit: string;
  detailsHtml: string;
}

interface IAuditGridProjection {
  groups?: IGroup[];
  items: IAuditGridItem[];
}

interface IAuditGroupData {
  detailsHtml: string;
  permissionItem: IPermissionAuditItem;
}

interface IPrincipalAccessMatch {
  matchItem: IPermissionAuditItem;
  scopeItem: IPermissionAuditItem;
}

const auditOptions: ICheckboxListOption[] = [
  { key: 'ExpandGroups', label: strings.AuditOptionExpandGroups },
  { key: 'IncludeListsWithUniquePermissions', label: strings.AuditOptionIncludeListsWithUniquePermissions },
  { key: 'IncludeListItemsWithUniquePermissions', label: strings.AuditOptionIncludeListItemsWithUniquePermissions }
];

const auditColumns: IColumn[] = [
  {
    key: 'groupName',
    name: strings.AuditResultsDisplayNameColumn,
    fieldName: 'groupName',
    minWidth: 160,
    maxWidth: 280,
    isResizable: true
  },
  {
    key: 'principalType',
    name: strings.AuditResultsPrincipalTypeColumn,
    fieldName: 'principalType',
    minWidth: 120,
    maxWidth: 180,
    isResizable: true
  },
  {
    key: 'permissionLevels',
    name: strings.AuditResultsPermissionLevelsColumn,
    fieldName: 'permissionLevels',
    minWidth: 180,
    maxWidth: 300,
    isResizable: true
  },
  {
    key: 'groupPath',
    name: strings.AuditResultsGroupPathColumn,
    fieldName: 'groupPath',
    minWidth: 220,
    maxWidth: 420,
    isResizable: true
  },
  {
    key: 'owner',
    name: strings.AuditResultsOwnerColumn,
    fieldName: 'owner',
    minWidth: 140,
    maxWidth: 240,
    isResizable: true
  },
  {
    key: 'hiddenInUi',
    name: strings.AuditResultsHiddenInUiColumn,
    fieldName: 'hiddenInUi',
    minWidth: 80,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'membersCanEdit',
    name: strings.AuditResultsMembersCanEditColumn,
    fieldName: 'membersCanEdit',
    minWidth: 120,
    maxWidth: 150,
    isResizable: true
  }
];

const auditDetailsContentHeight: number = 180;
const auditResultsGridHeight: number = 420;
const groupedViewStorageValue: string = 'true';

const toDisplayText = (value: unknown): string => `${value ?? ''}`;

const encodeHtml = (value: unknown): string => {
  const stringValue: string = toDisplayText(value);

  return stringValue
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
};

const formatPermissionLevels = (permissionLevels: IPermissionAuditLevel[]): string => {
  if (permissionLevels.length === 0) {
    return strings.AuditGroupNoDirectPermissions;
  }

  return permissionLevels.map((permissionLevel) => permissionLevel.name).join(', ');
};

const toYesNo = (value: boolean): string => value ? strings.AuditGroupYes : strings.AuditGroupNo;

const getGroupType = (
  permissionItem: IPermissionAuditItem,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  switch (permissionItem.principalId) {
    case associatedGroupIds.ownerGroupId:
      return strings.AuditGroupTypeOwners;
    case associatedGroupIds.memberGroupId:
      return strings.AuditGroupTypeMembers;
    case associatedGroupIds.visitorGroupId:
      return strings.AuditGroupTypeVisitors;
    default:
      return strings.AuditGroupTypeSharePointGroup;
  }
};

const getPrincipalTypeLabel = (
  permissionItem: IPermissionAuditItem,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  if (permissionItem.principalType === 'SharePointGroup' && permissionItem.sourceType === 'DirectPermission') {
    return getGroupType(permissionItem, associatedGroupIds);
  }

  switch (permissionItem.principalType) {
    case 'DistributionList':
      return strings.AuditPrincipalTypeDistributionList;
    case 'List':
      return strings.AuditPrincipalTypeList;
    case 'ListItem':
      return strings.AuditPrincipalTypeListItem;
    case 'Microsoft365Group':
      return strings.AuditPrincipalTypeMicrosoft365Group;
    case 'SecurityGroup':
      return strings.AuditPrincipalTypeSecurityGroup;
    case 'SharePointGroup':
      return strings.AuditGroupTypeSharePointGroup;
    case 'Site':
      return strings.AuditPrincipalTypeSite;
    case 'User':
      return strings.AuditPrincipalTypeUser;
    default:
      return strings.AuditPrincipalTypeUnknown;
  }
};

const getGroupPath = (permissionItem: IPermissionAuditItem): string => {
  if (permissionItem.path.length <= 1) {
    return permissionItem.displayName;
  }

  return permissionItem.path.slice(0, -1).join(' > ');
};

const isSecurableObject = (permissionItem: IPermissionAuditItem): boolean =>
  permissionItem.principalType === 'Site' || permissionItem.principalType === 'List' || permissionItem.principalType === 'ListItem';

const isSearchablePrincipal = (permissionItem: IPermissionAuditItem): boolean =>
  !isSecurableObject(permissionItem);

const principalMatchesSearch = (permissionItem: IPermissionAuditItem, normalizedSearchText: string): boolean => {
  const searchableValues: string[] = [
    permissionItem.displayName,
    permissionItem.loginName,
    permissionItem.personDetails?.email,
    permissionItem.personDetails?.userPrincipalName,
    permissionItem.groupDetails?.aadObjectId
  ].map((value) => toDisplayText(value).toLowerCase());

  return searchableValues.some((value) => value.indexOf(normalizedSearchText) > -1);
};

const findNearestScope = (permissionItem: IPermissionAuditItem, ancestors: IPermissionAuditItem[]): IPermissionAuditItem =>
  isSecurableObject(permissionItem)
    ? permissionItem
    : ancestors.slice().reverse().find((ancestor) => isSecurableObject(ancestor)) || permissionItem;

const findPrincipalAccessMatches = (
  permissionItems: IPermissionAuditItem[],
  searchText: string,
  ancestors: IPermissionAuditItem[] = []
): IPrincipalAccessMatch[] => {
  const normalizedSearchText: string = searchText.trim().toLowerCase();

  if (!normalizedSearchText) {
    return [];
  }

  return permissionItems.reduce((matches: IPrincipalAccessMatch[], permissionItem) => {
    if (isSearchablePrincipal(permissionItem) && principalMatchesSearch(permissionItem, normalizedSearchText)) {
      matches.push({
        matchItem: permissionItem,
        scopeItem: findNearestScope(permissionItem, ancestors)
      });
    }

    if (permissionItem.children?.length) {
      matches.push(...findPrincipalAccessMatches(permissionItem.children, searchText, [...ancestors, permissionItem]));
    }

    return matches;
  }, []);
};

const buildDetailsHtml = (
  permissionItem: IPermissionAuditItem,
  permissionLevels: string,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  const groupDetails = permissionItem.groupDetails;
  const membershipDetails: string[] = [
    `${strings.AuditGroupDetailsMembersCanEditLabel}: ${toYesNo(!!groupDetails?.allowMembersEditMembership)}`,
    `${strings.AuditGroupDetailsMembersOnlyLabel}: ${toYesNo(!!groupDetails?.onlyAllowMembersViewMembership)}`,
    `${strings.AuditGroupDetailsRequestsLabel}: ${toYesNo(!!groupDetails?.allowRequestToJoinLeave)}`,
    `${strings.AuditGroupDetailsRequestEmailLabel}: ${toDisplayText(groupDetails?.requestToJoinLeaveEmail)}`
  ];

  return `
    <h3>${encodeHtml(strings.AuditGroupDetailsHeading)}</h3>
    <dl>
      <dt>${encodeHtml(strings.AuditResultsPrincipalTypeColumn)}</dt>
      <dd>${encodeHtml(getPrincipalTypeLabel(permissionItem, associatedGroupIds))}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsPathLabel)}</dt>
      <dd>${encodeHtml(permissionItem.path.join(' > '))}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsIdLabel)}</dt>
      <dd>${encodeHtml(permissionItem.principalId)}</dd>
      <dt>${encodeHtml(strings.AuditResultsGroupNameColumn)}</dt>
      <dd>${encodeHtml(permissionItem.displayName)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsExpansionErrorLabel)}</dt>
      <dd>${encodeHtml(groupDetails?.expansionError)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsLoginNameLabel)}</dt>
      <dd>${encodeHtml(permissionItem.loginName)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsDescriptionLabel)}</dt>
      <dd>${encodeHtml(permissionItem.description)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsOwnerLabel)}</dt>
      <dd>${encodeHtml(groupDetails?.ownerTitle)}</dd>
      <dt>${encodeHtml(strings.AuditObjectDetailsUrlLabel)}</dt>
      <dd>${encodeHtml(permissionItem.objectDetails?.url || permissionItem.objectDetails?.serverRelativeUrl)}</dd>
      <dt>${encodeHtml(strings.AuditObjectDetailsItemCountLabel)}</dt>
      <dd>${encodeHtml(permissionItem.objectDetails?.itemCount)}</dd>
      <dt>${encodeHtml(strings.AuditObjectDetailsItemIdLabel)}</dt>
      <dd>${encodeHtml(permissionItem.objectDetails?.itemId)}</dd>
      <dt>${encodeHtml(strings.AuditResultsHiddenInUiColumn)}</dt>
      <dd>${encodeHtml(permissionItem.objectDetails?.hidden === undefined ? '' : toYesNo(permissionItem.objectDetails.hidden))}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsEmailLabel)}</dt>
      <dd>${encodeHtml(permissionItem.personDetails?.email)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsUserPrincipalNameLabel)}</dt>
      <dd>${encodeHtml(permissionItem.personDetails?.userPrincipalName)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsPermissionsLabel)}</dt>
      <dd>${encodeHtml(permissionLevels)}</dd>
      <dt>${encodeHtml(strings.AuditGroupDetailsMembershipLabel)}</dt>
      <dd>${encodeHtml(membershipDetails.join(' | '))}</dd>
    </dl>
  `;
};

const buildPrincipalSearchDetailsHtml = (
  auditResult: ICurrentSitePermissionGroupsResult,
  searchText: string,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  const matches: IPrincipalAccessMatch[] = findPrincipalAccessMatches(auditResult.groups, searchText);

  if (matches.length === 0) {
    return `
      <h3>${encodeHtml(strings.AuditPrincipalSearchResultsHeading)}</h3>
      <p>${encodeHtml(strings.AuditPrincipalSearchNoResults)}</p>
    `;
  }

  const rows: string = matches.map((match) => {
    const matchTypeLabel: string = getPrincipalTypeLabel(match.matchItem, associatedGroupIds);
    const scopeTypeLabel: string = getPrincipalTypeLabel(match.scopeItem, associatedGroupIds);

    return `
      <tr>
        <td>${encodeHtml(match.matchItem.displayName)}</td>
        <td>${encodeHtml(matchTypeLabel)}</td>
        <td>${encodeHtml(formatPermissionLevels(match.matchItem.permissionLevels))}</td>
        <td>${encodeHtml(`${match.scopeItem.displayName} (${scopeTypeLabel})`)}</td>
        <td>${encodeHtml(match.matchItem.path.join(' > '))}</td>
      </tr>
    `;
  }).join('');

  return `
    <h3>${encodeHtml(strings.AuditPrincipalSearchResultsHeading)}</h3>
    <table>
      <thead>
        <tr>
          <th>${encodeHtml(strings.AuditResultsDisplayNameColumn)}</th>
          <th>${encodeHtml(strings.AuditResultsPrincipalTypeColumn)}</th>
          <th>${encodeHtml(strings.AuditResultsPermissionLevelsColumn)}</th>
          <th>${encodeHtml(strings.AuditPrincipalSearchScopeColumn)}</th>
          <th>${encodeHtml(strings.AuditPrincipalSearchAccessViaColumn)}</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
};

const toAuditGridItem = (
  permissionItem: IPermissionAuditItem,
  associatedGroupIds: IAssociatedSharePointGroupIds
): IAuditGridItem => {
  const permissionLevelsText: string = formatPermissionLevels(permissionItem.permissionLevels);

  return {
    key: permissionItem.key,
    permissionItem,
    groupName: permissionItem.displayName,
    principalType: getPrincipalTypeLabel(permissionItem, associatedGroupIds),
    permissionLevels: permissionLevelsText,
    groupPath: getGroupPath(permissionItem),
    owner: toDisplayText(permissionItem.groupDetails?.ownerTitle),
    hiddenInUi: toYesNo(!!(permissionItem.groupDetails?.isHiddenInUi || permissionItem.objectDetails?.hidden)),
    membersCanEdit: toYesNo(!!permissionItem.groupDetails?.allowMembersEditMembership),
    detailsHtml: buildDetailsHtml(permissionItem, permissionLevelsText, associatedGroupIds)
  };
};

const flattenPermissionItems = (
  permissionItems: IPermissionAuditItem[],
  associatedGroupIds: IAssociatedSharePointGroupIds
): IAuditGridItem[] => permissionItems.reduce((items: IAuditGridItem[], permissionItem) => {
  items.push(toAuditGridItem(permissionItem, associatedGroupIds));

  if (permissionItem.children?.length) {
    items.push(...flattenPermissionItems(permissionItem.children, associatedGroupIds));
  }

  return items;
}, []);

const isGroupNode = (permissionItem: IPermissionAuditItem): boolean =>
  permissionItem.principalType !== 'User' || !!permissionItem.children?.length;

const toGroupHeaderName = (
  permissionItem: IPermissionAuditItem,
  associatedGroupIds: IAssociatedSharePointGroupIds,
  userCount: number
): string => {
  const principalTypeLabel: string = getPrincipalTypeLabel(permissionItem, associatedGroupIds);
  const permissionLevels: string = formatPermissionLevels(permissionItem.permissionLevels);
  const userCountLabel: string = userCount === 1 ? strings.AuditUserCountSingular : strings.AuditUserCountPlural;

  if (permissionItem.principalType === 'Site' || permissionItem.principalType === 'List' || permissionItem.principalType === 'ListItem') {
    return `${permissionItem.displayName} (${principalTypeLabel}) - ${userCount} ${userCountLabel}`;
  }

  return `${permissionItem.displayName} (${principalTypeLabel}; ${permissionLevels}) - ${userCount} ${userCountLabel}`;
};

const countVisibleUsers = (permissionItems: IPermissionAuditItem[]): number =>
  permissionItems.reduce((count, permissionItem) => {
    const childUserCount: number = permissionItem.children ? countVisibleUsers(permissionItem.children) : 0;

    return count + (!isGroupNode(permissionItem) && permissionItem.principalType === 'User' ? 1 : 0) + childUserCount;
  }, 0);

const toUserCountLabel = (userCount: number): string =>
  userCount === 1 ? strings.AuditUserCountSingular : strings.AuditUserCountPlural;

const buildGroupedPermissionItems = (
  permissionItems: IPermissionAuditItem[],
  associatedGroupIds: IAssociatedSharePointGroupIds
): IAuditGridProjection => {
  const items: IAuditGridItem[] = [];

  const buildGroups = (childPermissionItems: IPermissionAuditItem[], level: number, parentKey: string): IGroup[] => {
    const groups: IGroup[] = [];
    const hasChildGroups: boolean = childPermissionItems.some((permissionItem) => isGroupNode(permissionItem));
    const directPermissionItems: IPermissionAuditItem[] = childPermissionItems.filter((permissionItem) => !isGroupNode(permissionItem));

    if (hasChildGroups && directPermissionItems.length > 0) {
      const startIndex: number = items.length;
      directPermissionItems.forEach((permissionItem) => {
        items.push(toAuditGridItem(permissionItem, associatedGroupIds));
      });

      groups.push({
        key: `${parentKey}-direct-permissions`,
        name: `${strings.AuditDirectPermissionsGroupLabel} - ${directPermissionItems.length} ${toUserCountLabel(directPermissionItems.length)}`,
        startIndex,
        count: items.length - startIndex,
        isCollapsed: false,
        level
      });
    }

    childPermissionItems.forEach((permissionItem) => {
      if (!isGroupNode(permissionItem)) {
        if (!hasChildGroups) {
          items.push(toAuditGridItem(permissionItem, associatedGroupIds));
        }

        return;
      }

      const startIndex: number = items.length;
      const childGroups: IGroup[] = buildGroups(permissionItem.children || [], level + 1, permissionItem.key);
      const count: number = items.length - startIndex;
      const permissionLevelsText: string = formatPermissionLevels(permissionItem.permissionLevels);

      groups.push({
        key: permissionItem.key,
        name: toGroupHeaderName(permissionItem, associatedGroupIds, countVisibleUsers(permissionItem.children || [])),
        startIndex,
        count,
        children: childGroups,
        data: {
          detailsHtml: buildDetailsHtml(permissionItem, permissionLevelsText, associatedGroupIds),
          permissionItem
        } as IAuditGroupData,
        isCollapsed: false,
        level
      });
    });

    return groups;
  };

  return {
    groups: buildGroups(permissionItems, 0, 'root'),
    items
  };
};

const readGroupedViewPreference = (preferenceKey: string): boolean => {
  try {
    return window.localStorage.getItem(preferenceKey) !== 'false';
  } catch {
    return true;
  }
};

const writeGroupedViewPreference = (preferenceKey: string, isGroupedView: boolean): void => {
  try {
    window.localStorage.setItem(preferenceKey, isGroupedView ? groupedViewStorageValue : 'false');
  } catch {
    // Storage can be unavailable in privacy-restricted browser contexts.
  }
};

const projectAuditResult = (
  auditResult: ICurrentSitePermissionGroupsResult,
  isGroupedView: boolean
): IAuditGridProjection => isGroupedView
  ? buildGroupedPermissionItems(auditResult.groups, auditResult.associatedGroupIds)
  : {
    groups: undefined,
    items: flattenPermissionItems(auditResult.groups, auditResult.associatedGroupIds)
  };

const toCsvValue = (value: unknown): string => {
  const stringValue: string = toDisplayText(value);

  if (/[",\r\n]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }

  return stringValue;
};

const toCsvHeader = (localizedHeader: string | undefined, fallbackHeader: string): string =>
  localizedHeader || fallbackHeader;

const buildCsvContent = (auditResult: ICurrentSitePermissionGroupsResult): string => {
  const rows: IAuditGridItem[] = flattenPermissionItems(auditResult.groups, auditResult.associatedGroupIds);
  const headers: string[] = [
    toCsvHeader(strings.AuditExportKeyColumn, 'Key'),
    toCsvHeader(strings.AuditResultsDisplayNameColumn, 'Display Name'),
    toCsvHeader(strings.AuditResultsPrincipalTypeColumn, 'Type'),
    toCsvHeader(strings.AuditExportSourceTypeColumn, 'Source type'),
    toCsvHeader(strings.AuditResultsPermissionLevelsColumn, 'Permission levels'),
    toCsvHeader(strings.AuditGroupDetailsPathLabel, 'Path'),
    toCsvHeader(strings.AuditResultsGroupPathColumn, 'Group path'),
    toCsvHeader(strings.AuditExportDepthColumn, 'Depth'),
    toCsvHeader(strings.AuditExportParentKeyColumn, 'Parent key'),
    toCsvHeader(strings.AuditGroupDetailsIdLabel, 'Id'),
    toCsvHeader(strings.AuditGroupDetailsLoginNameLabel, 'Login name'),
    toCsvHeader(strings.AuditGroupDetailsDescriptionLabel, 'Description'),
    toCsvHeader(strings.AuditGroupDetailsOwnerLabel, 'Owner'),
    toCsvHeader(strings.AuditResultsHiddenInUiColumn, 'Hidden'),
    toCsvHeader(strings.AuditGroupDetailsMembersCanEditLabel, 'Members can edit membership'),
    toCsvHeader(strings.AuditGroupDetailsMembersOnlyLabel, 'Only members can view membership'),
    toCsvHeader(strings.AuditGroupDetailsRequestsLabel, 'Join and leave requests'),
    toCsvHeader(strings.AuditGroupDetailsRequestEmailLabel, 'Request email'),
    toCsvHeader(strings.AuditGroupDetailsExpansionErrorLabel, 'Expansion error'),
    toCsvHeader(strings.AuditExportAadObjectIdColumn, 'Entra object id'),
    toCsvHeader(strings.AuditObjectDetailsUrlLabel, 'Url'),
    toCsvHeader(strings.AuditExportBaseTemplateColumn, 'Base template'),
    toCsvHeader(strings.AuditObjectDetailsItemCountLabel, 'Item count'),
    toCsvHeader(strings.AuditObjectDetailsItemIdLabel, 'Item id'),
    toCsvHeader(strings.AuditGroupDetailsEmailLabel, 'Email'),
    toCsvHeader(strings.AuditGroupDetailsUserPrincipalNameLabel, 'User principal name'),
    toCsvHeader(strings.AuditExportJobTitleColumn, 'Job title'),
    toCsvHeader(strings.AuditExportDepartmentColumn, 'Department'),
    toCsvHeader(strings.AuditExportAdditionalDetailsColumn, 'Additional details')
  ];
  const csvRows: string[] = rows.map((row) => {
    const permissionItem: IPermissionAuditItem = row.permissionItem;
    const groupDetails = permissionItem.groupDetails;
    const objectDetails = permissionItem.objectDetails;
    const personDetails = permissionItem.personDetails;
    const values: unknown[] = [
      permissionItem.key,
      row.groupName,
      row.principalType,
      permissionItem.sourceType,
      row.permissionLevels,
      permissionItem.path.join(' > '),
      row.groupPath,
      permissionItem.depth,
      permissionItem.parentKey,
      permissionItem.principalId,
      permissionItem.loginName,
      permissionItem.description,
      groupDetails?.ownerTitle,
      row.hiddenInUi,
      groupDetails?.allowMembersEditMembership === undefined ? '' : toYesNo(groupDetails.allowMembersEditMembership),
      groupDetails?.onlyAllowMembersViewMembership === undefined ? '' : toYesNo(groupDetails.onlyAllowMembersViewMembership),
      groupDetails?.allowRequestToJoinLeave === undefined ? '' : toYesNo(groupDetails.allowRequestToJoinLeave),
      groupDetails?.requestToJoinLeaveEmail,
      groupDetails?.expansionError,
      groupDetails?.aadObjectId,
      objectDetails?.url || objectDetails?.serverRelativeUrl,
      objectDetails?.baseTemplate,
      objectDetails?.itemCount,
      objectDetails?.itemId,
      personDetails?.email,
      personDetails?.userPrincipalName,
      personDetails?.jobTitle,
      personDetails?.department,
      permissionItem.details?.map((detail) => `${detail.label}: ${detail.value}`).join(' | ')
    ];

    return values.map(toCsvValue).join(',');
  });

  return `\uFEFF${[headers.map(toCsvValue).join(','), ...csvRows].join('\r\n')}`;
};

const downloadCsv = (csvContent: string): void => {
  const blob: Blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url: string = URL.createObjectURL(blob);
  const link: HTMLAnchorElement = document.createElement('a');
  const timestamp: string = new Date().toISOString().replace(/[:.]/g, '-');

  link.href = url;
  link.download = `permission-audit-${timestamp}.csv`;
  link.click();
  URL.revokeObjectURL(url);
};

export const AuditPage: React.FunctionComponent<IAuditPageProps> = (props) => {
  const [selectedOptionKeys, setSelectedOptionKeys] = React.useState<string[]>([]);
  const [auditGroups, setAuditGroups] = React.useState<IGroup[] | undefined>(undefined);
  const [auditItems, setAuditItems] = React.useState<IAuditGridItem[]>([]);
  const [auditResult, setAuditResult] = React.useState<ICurrentSitePermissionGroupsResult | undefined>(undefined);
  const [selectedDetailsHtml, setSelectedDetailsHtml] = React.useState<string>('');
  const [principalSearchText, setPrincipalSearchText] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [isGroupedView, setIsGroupedView] = React.useState<boolean>(() => readGroupedViewPreference(props.groupedViewPreferenceKey));
  const isSharePointReady: boolean = !!props.sharePointPermissionAuditService;
  const selectionRef = React.useRef<Selection>();
  const includeListsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') > -1;
  const includeListItemsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') > -1;
  const includeListAuditScope: boolean = includeListsWithUniquePermissions || includeListItemsWithUniquePermissions;
  const auditOptionsWithDependencies: ICheckboxListOption[] = React.useMemo(() => [
    auditOptions[0],
    auditOptions[1],
    {
      key: 'IncludeHiddenListsWithUniquePermissions',
      label: strings.AuditOptionIncludeHiddenListsWithUniquePermissions,
      disabled: !includeListAuditScope
    },
    auditOptions[2]
  ], [includeListAuditScope]);

  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        const selectedItems: IAuditGridItem[] = selectionRef.current?.getSelection() as IAuditGridItem[];
        setSelectedDetailsHtml(selectedItems[0]?.detailsHtml || '');
      }
    });
  }

  const onStartAudit = React.useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setAuditResult(undefined);
    setAuditGroups(undefined);
    setAuditItems([]);
    setSelectedDetailsHtml('');

    try {
      const shouldExpandGroups: boolean = selectedOptionKeys.indexOf('ExpandGroups') > -1;
      const shouldIncludeListsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') > -1;
      const shouldIncludeHiddenListsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeHiddenListsWithUniquePermissions') > -1;
      const shouldIncludeListItemsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') > -1;

      const updateAuditResult = (auditResult: ICurrentSitePermissionGroupsResult): void => {
        const projection: IAuditGridProjection = projectAuditResult(auditResult, isGroupedView);

        setAuditResult(auditResult);
        setAuditGroups(projection.groups);
        setAuditItems(projection.items);
      };

      if (shouldIncludeListsWithUniquePermissions || shouldIncludeListItemsWithUniquePermissions) {
        const auditOptions: ISharePointPermissionAuditOptions = {
          expandGroups: shouldExpandGroups,
          includeHiddenListsWithUniquePermissions: shouldIncludeHiddenListsWithUniquePermissions,
          includeListItemsWithUniquePermissions: shouldIncludeListItemsWithUniquePermissions,
          includeListsWithUniquePermissions: shouldIncludeListsWithUniquePermissions,
          onAuditUpdated: updateAuditResult
        };
        const auditResult: ICurrentSitePermissionGroupsResult =
          await props.sharePointPermissionAuditService.getPermissionAuditAsync(auditOptions);

        updateAuditResult(auditResult);
      } else if (selectedOptionKeys.length === 0) {
        const siteGroupsResult: ICurrentSitePermissionGroupsResult =
          await props.sharePointPermissionAuditService.getCurrentSitePermissionGroupsWithMetadataAsync();
        const projection: IAuditGridProjection = projectAuditResult(siteGroupsResult, isGroupedView);

        setAuditResult(siteGroupsResult);
        setAuditGroups(projection.groups);
        setAuditItems(projection.items);
      } else if (shouldExpandGroups) {
        const siteGroupsResult: ICurrentSitePermissionGroupsResult =
          await props.sharePointPermissionAuditService.getCurrentSitePermissionGroupsWithMetadataAsync(true);
        const projection: IAuditGridProjection = projectAuditResult(siteGroupsResult, isGroupedView);

        setAuditResult(siteGroupsResult);
        setAuditGroups(projection.groups);
        setAuditItems(projection.items);
      }
    } catch (error) {
      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
    } finally {
      setIsLoading(false);
    }
  }, [isGroupedView, props.sharePointPermissionAuditService, selectedOptionKeys]);

  const onStartButtonClick = React.useCallback((): void => {
    onStartAudit().catch((error) => {
      setIsLoading(false);
      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
    });
  }, [onStartAudit]);

  const onAuditOptionsChange = React.useCallback((nextSelectedOptionKeys: string[]): void => {
    if (
      nextSelectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') === -1 &&
      nextSelectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') === -1
    ) {
      setSelectedOptionKeys(nextSelectedOptionKeys.filter((key) => key !== 'IncludeHiddenListsWithUniquePermissions'));
      return;
    }

    setSelectedOptionKeys(nextSelectedOptionKeys);
  }, []);

  const onGroupedViewChange = React.useCallback((ev?: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    const nextIsGroupedView: boolean = !!checked;

    setIsGroupedView(nextIsGroupedView);
    writeGroupedViewPreference(props.groupedViewPreferenceKey, nextIsGroupedView);

    if (auditResult) {
      const projection: IAuditGridProjection = projectAuditResult(auditResult, nextIsGroupedView);

      setAuditGroups(projection.groups);
      setAuditItems(projection.items);
    }
  }, [auditResult, props.groupedViewPreferenceKey]);

  const onExportCsvClick = React.useCallback((): void => {
    if (!auditResult) {
      return;
    }

    downloadCsv(buildCsvContent(auditResult));
  }, [auditResult]);

  const onPrincipalSearch = React.useCallback((): void => {
    if (!auditResult) {
      return;
    }

    setSelectedDetailsHtml(buildPrincipalSearchDetailsHtml(
      auditResult,
      principalSearchText,
      auditResult.associatedGroupIds
    ));
  }, [auditResult, principalSearchText]);

  const onPrincipalSearchClear = React.useCallback((): void => {
    setPrincipalSearchText('');
    setSelectedDetailsHtml('');
  }, []);

  const onRenderGroupHeader = React.useCallback((
    groupHeaderProps?: IGroupHeaderProps,
    defaultRender?: (props?: IGroupHeaderProps) => JSX.Element | null
  ): JSX.Element | null => {
    if (!groupHeaderProps || !defaultRender) {
      return null;
    }

    const onClick = (): void => {
      const groupData: IAuditGroupData | undefined = groupHeaderProps.group?.data as IAuditGroupData | undefined;
      setSelectedDetailsHtml(groupData?.detailsHtml || '');
    };

    return (
      <div onClick={onClick}>
        {defaultRender(groupHeaderProps)}
      </div>
    );
  }, []);

  const onRenderDetailsHeader = React.useCallback((
    detailsHeaderProps?: IDetailsHeaderProps,
    defaultRender?: (props?: IDetailsHeaderProps) => JSX.Element | null
  ): JSX.Element | null => {
    if (!detailsHeaderProps || !defaultRender) {
      return null;
    }

    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender(detailsHeaderProps)}
      </Sticky>
    );
  }, []);

  return (
    <section className={styles.auditPage}>
      <div className={styles.header}>
        <h2 className={styles.title}>{strings.AuditPageTitle}</h2>
        <p className={styles.subtitle}>{strings.AuditPageSubtitle}</p>
      </div>

      <CheckboxList
        label={strings.AuditOptionsLabel}
        options={auditOptionsWithDependencies}
        selectedKeys={selectedOptionKeys}
        onSelectionChange={onAuditOptionsChange}
      />

      <Stack horizontal horizontalAlign="start" className={styles.commandBar}>
        <PrimaryButton
          text={strings.StartButtonLabel}
          disabled={!isSharePointReady || isLoading}
          onClick={onStartButtonClick}
        />
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          text={strings.ExportCsvButtonLabel}
          disabled={!auditResult || isLoading}
          onClick={onExportCsvClick}
        />
        {isLoading && (
          <Spinner
            label={strings.AuditLoadingLabel}
            size={SpinnerSize.small}
          />
        )}
      </Stack>

      <div className={styles.gridOptions}>
        <Toggle
          checked={isGroupedView}
          inlineLabel
          label={strings.AuditGroupedViewToggleLabel}
          onChange={onGroupedViewChange}
        />
      </div>

      <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.searchBar}>
        <SearchBox
          ariaLabel={strings.AuditPrincipalSearchLabel}
          disabled={!auditResult || isLoading}
          placeholder={strings.AuditPrincipalSearchPlaceholder}
          value={principalSearchText}
          onChange={(_, nextValue) => setPrincipalSearchText(nextValue || '')}
          onSearch={onPrincipalSearch}
        />
        <DefaultButton
          text={strings.AuditPrincipalSearchButtonLabel}
          disabled={!auditResult || isLoading || principalSearchText.trim().length === 0}
          onClick={onPrincipalSearch}
        />
        <DefaultButton
          text={strings.AuditPrincipalSearchClearButtonLabel}
          disabled={principalSearchText.length === 0}
          onClick={onPrincipalSearchClear}
        />
      </Stack>

      <div className={styles.grid} style={{ height: auditResultsGridHeight }}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.always}>
          <DetailsList
            items={auditItems}
            columns={auditColumns}
            constrainMode={ConstrainMode.unconstrained}
            groups={auditGroups}
            groupProps={{
              headerProps: {
                styles: {
                  headerCount: {
                    display: 'none'
                  }
                }
              },
              onRenderHeader: onRenderGroupHeader,
              showEmptyGroups: true
            }}
            selection={selectionRef.current}
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.justified}
            onRenderDetailsHeader={onRenderDetailsHeader}
            ariaLabelForGrid={strings.AuditResultsGridAriaLabel}
          />
        </ScrollablePane>
      </div>

      <HtmlContentDisplay
        ariaLabel={strings.AuditDetailsContentAriaLabel}
        className={styles.detailsContent}
        height={auditDetailsContentHeight}
        html={selectedDetailsHtml}
      />
    </section>
  );
};
