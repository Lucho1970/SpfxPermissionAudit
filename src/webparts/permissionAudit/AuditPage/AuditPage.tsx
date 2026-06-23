import * as React from 'react';
import { ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, IDetailsHeaderProps, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { IGroup, IGroupHeaderProps } from '@fluentui/react/lib/GroupedList';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import { ScrollablePane, ScrollbarVisibility } from '@fluentui/react/lib/ScrollablePane';
import { Selection } from '@fluentui/react/lib/Selection';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Sticky, StickyPositionType } from '@fluentui/react/lib/Sticky';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { CheckboxList, ICheckboxListOption } from '../../../common/controls/CheckboxList';
import { HtmlContentDisplay } from '../../../common/controls/HtmlContentDisplay';
import type { IAuditPageProps } from './IAuditPageProps';
import type { IPermissionAuditItem, IPermissionAuditLevel } from '../models';
import type {
  IAssociatedSharePointGroupIds,
  ICurrentSitePermissionGroupsResult,
  IPrincipalAccessSearchResult
} from '../services';
import type { IDirectoryPersonInfo } from '../services';
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

interface IPrincipalSuggestionPersona extends IPersonaProps {
  directoryPerson: IDirectoryPersonInfo;
}

interface ILoadedListLayer {
  expandedNodes?: IPermissionAuditItem[];
  includeHidden: boolean;
  nodes: IPermissionAuditItem[];
}

interface IAuditDataCache {
  baseResult?: ICurrentSitePermissionGroupsResult;
  expandedBaseResult?: ICurrentSitePermissionGroupsResult;
  listItemsLayer?: ILoadedListLayer;
  listPermissionsLayer?: ILoadedListLayer;
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
    isResizable: true,
    onRender: (item?: IAuditGridItem): JSX.Element | string => item?.permissionItem.principalType === 'LoadMore'
      ? <span className={(styles as unknown as { [key: string]: string }).loadMoreLink}>{strings.AuditLoadRemainingGroupMembersLabel}</span>
      : `${item?.groupName ?? ''}`
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

const groupedViewStorageValue: string = 'true';
const auditAutoRunDebounceMs: number = 400;
const auditSplitPaneHeight: number = 612;
const auditResizeHandleHeight: number = 12;
const auditDefaultGridHeight: number = 420;
const auditMinimumGridHeight: number = 220;
const auditMinimumDetailsHeight: number = 140;

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

const toDetailValue = (value: unknown): string => {
  const text: string = toDisplayText(value).trim();

  return text || strings.AuditDetailsEmptyValue;
};

const joinPath = (path: string[]): string => path.join(strings.AuditDetailsPathValueSeparator);

const clamp = (value: number, minValue: number, maxValue: number): number =>
  Math.min(Math.max(value, minValue), maxValue);

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
    case 'LoadMore':
      return strings.AuditPrincipalTypeLoadMore;
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

const getDisplayName = (permissionItem: IPermissionAuditItem): string =>
  permissionItem.principalType === 'LoadMore'
    ? strings.AuditLoadRemainingGroupMembersLabel
    : permissionItem.displayName;

const getGroupPath = (permissionItem: IPermissionAuditItem): string => {
  if (permissionItem.path.length <= 1) {
    return permissionItem.displayName;
  }

  return permissionItem.path.slice(0, -1).join(' > ');
};

const buildDetailsHtml = (
  permissionItem: IPermissionAuditItem,
  permissionLevels: string,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  const groupDetails = permissionItem.groupDetails;
  const objectDetails = permissionItem.objectDetails;
  const personDetails = permissionItem.personDetails;
  const tags: string[] = [
    getPrincipalTypeLabel(permissionItem, associatedGroupIds),
    permissionItem.sourceType
  ];

  if (permissionLevels !== strings.AuditGroupNoDirectPermissions) {
    tags.push(permissionLevels);
  }

  const renderDetailRow = (label: string, value: unknown): string => `
    <div class="audit-detail-row">
      <dt>${encodeHtml(label)}</dt>
      <dd>${encodeHtml(toDetailValue(value))}</dd>
    </div>
  `;

  const renderDetailSection = (title: string, rows: string[]): string => `
    <section class="audit-detail-section">
      <h4>${encodeHtml(title)}</h4>
      <dl>${rows.join('')}</dl>
    </section>
  `;

  const renderTagList = (tagValues: string[]): string => `
    <div class="audit-detail-tags" aria-label="${encodeHtml(strings.AuditDetailsTagsLabel)}">
      ${tagValues.map((tagValue) => `<span class="audit-detail-tag">${encodeHtml(tagValue)}</span>`).join('')}
    </div>
  `;

  const renderAdditionalDetails = (): string => {
    if (!permissionItem.details?.length) {
      return '';
    }

    return renderDetailSection(
      strings.AuditDetailsAdditionalSection,
      permissionItem.details.map((detail) => renderDetailRow(detail.label, detail.value))
    );
  };

  const renderPrincipalTemplate = (objectTypeLabel: string, extraSections: string[]): string => `
    <div class="audit-detail-card">
      <header class="audit-detail-header">
        <div>
          <p class="audit-detail-eyebrow">${encodeHtml(objectTypeLabel)}</p>
          <h3>${encodeHtml(permissionItem.displayName)}</h3>
          <p class="audit-detail-path">${encodeHtml(joinPath(permissionItem.path))}</p>
        </div>
        ${renderTagList(tags)}
      </header>
      ${renderDetailSection(strings.AuditDetailsOverviewSection, [
        renderDetailRow(strings.AuditResultsDisplayNameColumn, permissionItem.displayName),
        renderDetailRow(strings.AuditResultsPrincipalTypeColumn, getPrincipalTypeLabel(permissionItem, associatedGroupIds)),
        renderDetailRow(strings.AuditDetailsSourceTypeLabel, permissionItem.sourceType),
        renderDetailRow(strings.AuditGroupDetailsPathLabel, joinPath(permissionItem.path)),
        renderDetailRow(strings.AuditDetailsInheritedFromLabel, permissionItem.path.length > 1 ? permissionItem.path.slice(0, -1).join(strings.AuditDetailsPathValueSeparator) : ''),
        renderDetailRow(strings.AuditDetailsDepthLabel, permissionItem.depth),
        renderDetailRow(strings.AuditDetailsParentKeyLabel, permissionItem.parentKey),
        renderDetailRow(strings.AuditDetailsDirectChildrenLabel, permissionItem.children?.length || 0)
      ])}
      ${renderDetailSection(strings.AuditDetailsAccessSection, [
        renderDetailRow(strings.AuditGroupDetailsPermissionsLabel, permissionLevels),
        renderDetailRow(strings.AuditGroupDetailsExpansionErrorLabel, groupDetails?.expansionError)
      ])}
      ${extraSections.join('')}
      ${renderAdditionalDetails()}
    </div>
  `;

  switch (permissionItem.principalType) {
    case 'Site':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeSite, [
        renderDetailSection(strings.AuditDetailsObjectSection, [
          renderDetailRow(strings.AuditGroupDetailsDescriptionLabel, permissionItem.description),
          renderDetailRow(strings.AuditObjectDetailsUrlLabel, objectDetails?.url || objectDetails?.serverRelativeUrl),
          renderDetailRow(strings.AuditDetailsDirectChildrenLabel, permissionItem.children?.length || 0)
        ])
      ]);
    case 'List':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeList, [
        renderDetailSection(strings.AuditDetailsObjectSection, [
          renderDetailRow(strings.AuditGroupDetailsDescriptionLabel, permissionItem.description),
          renderDetailRow(strings.AuditObjectDetailsUrlLabel, objectDetails?.serverRelativeUrl || objectDetails?.url),
          renderDetailRow(strings.AuditExportBaseTemplateColumn, objectDetails?.baseTemplate),
          renderDetailRow(strings.AuditObjectDetailsItemCountLabel, objectDetails?.itemCount),
          renderDetailRow(strings.AuditResultsHiddenInUiColumn, objectDetails?.hidden === undefined ? '' : toYesNo(objectDetails.hidden))
        ])
      ]);
    case 'ListItem':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeListItem, [
        renderDetailSection(strings.AuditDetailsObjectSection, [
          renderDetailRow(strings.AuditObjectDetailsItemIdLabel, objectDetails?.itemId),
          renderDetailRow(strings.AuditObjectDetailsUrlLabel, objectDetails?.serverRelativeUrl || objectDetails?.url),
          renderDetailRow(strings.AuditDetailsDirectChildrenLabel, permissionItem.children?.length || 0)
        ])
      ]);
    case 'User':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeUser, [
        renderDetailSection(strings.AuditDetailsIdentitySection, [
          renderDetailRow(strings.AuditGroupDetailsIdLabel, permissionItem.principalId),
          renderDetailRow(strings.AuditGroupDetailsLoginNameLabel, permissionItem.loginName),
          renderDetailRow(strings.AuditGroupDetailsEmailLabel, personDetails?.email),
          renderDetailRow(strings.AuditGroupDetailsUserPrincipalNameLabel, personDetails?.userPrincipalName),
          renderDetailRow(strings.AuditExportJobTitleColumn, personDetails?.jobTitle),
          renderDetailRow(strings.AuditExportDepartmentColumn, personDetails?.department)
        ])
      ]);
    case 'SharePointGroup':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeSharePointGroup, [
        renderDetailSection(strings.AuditDetailsIdentitySection, [
          renderDetailRow(strings.AuditGroupDetailsIdLabel, permissionItem.principalId),
          renderDetailRow(strings.AuditGroupDetailsLoginNameLabel, permissionItem.loginName),
          renderDetailRow(strings.AuditGroupDetailsDescriptionLabel, permissionItem.description),
          renderDetailRow(strings.AuditGroupDetailsOwnerLabel, groupDetails?.ownerTitle)
        ]),
        renderDetailSection(strings.AuditDetailsMembershipSection, [
          renderDetailRow(strings.AuditResultsMembersCanEditColumn, groupDetails?.allowMembersEditMembership === undefined ? '' : toYesNo(groupDetails.allowMembersEditMembership)),
          renderDetailRow(strings.AuditGroupDetailsMembersOnlyLabel, groupDetails?.onlyAllowMembersViewMembership === undefined ? '' : toYesNo(groupDetails.onlyAllowMembersViewMembership)),
          renderDetailRow(strings.AuditGroupDetailsRequestsLabel, groupDetails?.allowRequestToJoinLeave === undefined ? '' : toYesNo(groupDetails.allowRequestToJoinLeave)),
          renderDetailRow(strings.AuditGroupDetailsRequestEmailLabel, groupDetails?.requestToJoinLeaveEmail),
          renderDetailRow(strings.AuditResultsHiddenInUiColumn, groupDetails?.isHiddenInUi === undefined ? '' : toYesNo(groupDetails.isHiddenInUi)),
          renderDetailRow(strings.AuditDetailsDirectChildrenLabel, permissionItem.children?.length || 0)
        ])
      ]);
    case 'SecurityGroup':
    case 'DistributionList':
    case 'Microsoft365Group':
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypeDirectoryGroup, [
        renderDetailSection(strings.AuditDetailsDirectorySection, [
          renderDetailRow(strings.AuditGroupDetailsIdLabel, permissionItem.principalId),
          renderDetailRow(strings.AuditGroupDetailsLoginNameLabel, permissionItem.loginName),
          renderDetailRow(strings.AuditExportAadObjectIdColumn, groupDetails?.aadObjectId),
          renderDetailRow(strings.AuditResultsHiddenInUiColumn, groupDetails?.isHiddenInUi === undefined ? '' : toYesNo(groupDetails.isHiddenInUi)),
          renderDetailRow(strings.AuditDetailsDirectChildrenLabel, permissionItem.children?.length || 0)
        ])
      ]);
    default:
      return renderPrincipalTemplate(strings.AuditDetailsObjectTypePrincipal, [
        renderDetailSection(strings.AuditDetailsIdentitySection, [
          renderDetailRow(strings.AuditGroupDetailsIdLabel, permissionItem.principalId),
          renderDetailRow(strings.AuditGroupDetailsLoginNameLabel, permissionItem.loginName),
          renderDetailRow(strings.AuditGroupDetailsDescriptionLabel, permissionItem.description),
          renderDetailRow(strings.AuditExportAadObjectIdColumn, groupDetails?.aadObjectId)
        ])
      ]);
  }
};

const buildPrincipalSearchDetailsHtml = (
  searchResult: IPrincipalAccessSearchResult,
  associatedGroupIds: IAssociatedSharePointGroupIds
): string => {
  const { principal, matches } = searchResult;

  if (matches.length === 0) {
    return `
      <div class="audit-detail-card">
        <header class="audit-detail-header">
          <div>
            <p class="audit-detail-eyebrow">${encodeHtml(strings.AuditPrincipalTypeUser)}</p>
            <h3>${encodeHtml(principal.displayName)}</h3>
            <p class="audit-detail-path">${encodeHtml(toDetailValue(principal.userPrincipalName || principal.email))}</p>
          </div>
        </header>
        <section class="audit-detail-section">
          <h4>${encodeHtml(strings.AuditPrincipalSearchResultsHeading)}</h4>
          <p>${encodeHtml(strings.AuditPrincipalSearchNoResults)}</p>
        </section>
      </div>
    `;
  }

  const uniqueScopeCount: number = matches.filter((match) =>
    match.scopeItem.principalType === 'List' || match.scopeItem.principalType === 'ListItem'
  ).length;
  const rows: string = matches.map((match) => {
    const matchTypeLabel: string = getPrincipalTypeLabel(match.matchItem, associatedGroupIds);
    const scopeTypeLabel: string = getPrincipalTypeLabel(match.scopeItem, associatedGroupIds);

    return `
      <tr>
        <td>${encodeHtml(match.matchItem.displayName)}</td>
        <td>${encodeHtml(matchTypeLabel)}</td>
        <td>${encodeHtml(formatPermissionLevels(match.matchItem.permissionLevels))}</td>
        <td>${encodeHtml(`${match.scopeItem.displayName} (${scopeTypeLabel})`)}</td>
        <td>${encodeHtml(joinPath(match.matchItem.path))}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="audit-detail-card">
      <header class="audit-detail-header">
        <div>
          <p class="audit-detail-eyebrow">${encodeHtml(strings.AuditPrincipalTypeUser)}</p>
          <h3>${encodeHtml(principal.displayName)}</h3>
          <p class="audit-detail-path">${encodeHtml(toDetailValue(principal.userPrincipalName || principal.email))}</p>
        </div>
        <div class="audit-detail-tags" aria-label="${encodeHtml(strings.AuditDetailsTagsLabel)}">
          <span class="audit-detail-tag">${encodeHtml(`${matches.length} ${strings.AuditPrincipalSearchResultsHeading}`)}</span>
          <span class="audit-detail-tag">${encodeHtml(`${uniqueScopeCount} ${strings.AuditPrincipalSearchScopeColumn}`)}</span>
        </div>
      </header>
      <section class="audit-detail-section">
        <h4>${encodeHtml(strings.AuditDetailsIdentitySection)}</h4>
        <dl>
          <div class="audit-detail-row">
            <dt>${encodeHtml(strings.AuditResultsDisplayNameColumn)}</dt>
            <dd>${encodeHtml(toDetailValue(principal.displayName))}</dd>
          </div>
          <div class="audit-detail-row">
            <dt>${encodeHtml(strings.AuditGroupDetailsEmailLabel)}</dt>
            <dd>${encodeHtml(toDetailValue(principal.email))}</dd>
          </div>
          <div class="audit-detail-row">
            <dt>${encodeHtml(strings.AuditGroupDetailsUserPrincipalNameLabel)}</dt>
            <dd>${encodeHtml(toDetailValue(principal.userPrincipalName))}</dd>
          </div>
          <div class="audit-detail-row">
            <dt>${encodeHtml(strings.AuditExportDepartmentColumn)}</dt>
            <dd>${encodeHtml(toDetailValue(principal.department))}</dd>
          </div>
          <div class="audit-detail-row">
            <dt>${encodeHtml(strings.AuditExportJobTitleColumn)}</dt>
            <dd>${encodeHtml(toDetailValue(principal.jobTitle))}</dd>
          </div>
        </dl>
      </section>
      <section class="audit-detail-section">
        <h4>${encodeHtml(strings.AuditPrincipalSearchResultsHeading)}</h4>
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
      </section>
    </div>
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
    groupName: getDisplayName(permissionItem),
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
  (permissionItem.principalType !== 'User' && permissionItem.principalType !== 'LoadMore') || !!permissionItem.children?.length;

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

const readPanelHeightPreference = (preferenceKey: string): number => {
  try {
    const storedValue: string | null = window.localStorage.getItem(preferenceKey);
    const parsedValue: number = storedValue ? Number(storedValue) : auditDefaultGridHeight;

    return clamp(
      Number.isFinite(parsedValue) ? parsedValue : auditDefaultGridHeight,
      auditMinimumGridHeight,
      auditSplitPaneHeight - auditResizeHandleHeight - auditMinimumDetailsHeight
    );
  } catch {
    return auditDefaultGridHeight;
  }
};

const writePanelHeightPreference = (preferenceKey: string, height: number): void => {
  try {
    window.localStorage.setItem(preferenceKey, `${height}`);
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

const replacePermissionItem = (
  permissionItems: IPermissionAuditItem[],
  itemKey: string,
  replacementItems: IPermissionAuditItem[]
): IPermissionAuditItem[] => permissionItems.reduce((nextItems: IPermissionAuditItem[], permissionItem) => {
  if (permissionItem.key === itemKey) {
    nextItems.push(...replacementItems);
    return nextItems;
  }

  nextItems.push({
    ...permissionItem,
    children: permissionItem.children
      ? replacePermissionItem(permissionItem.children, itemKey, replacementItems)
      : permissionItem.children
  });

  return nextItems;
}, []);

const replacePermissionItemInResult = (
  auditResult: ICurrentSitePermissionGroupsResult | undefined,
  itemKey: string,
  replacementItems: IPermissionAuditItem[]
): ICurrentSitePermissionGroupsResult | undefined => auditResult
  ? {
    ...auditResult,
    groups: replacePermissionItem(auditResult.groups, itemKey, replacementItems)
  }
  : auditResult;

const replacePermissionItemInLayer = (
  listLayer: ILoadedListLayer | undefined,
  itemKey: string,
  replacementItems: IPermissionAuditItem[]
): ILoadedListLayer | undefined => listLayer
  ? {
    ...listLayer,
    expandedNodes: listLayer.expandedNodes
      ? replacePermissionItem(listLayer.expandedNodes, itemKey, replacementItems)
      : listLayer.expandedNodes,
    nodes: replacePermissionItem(listLayer.nodes, itemKey, replacementItems)
  }
  : listLayer;

const replacePermissionItemInCache = (
  cache: IAuditDataCache,
  itemKey: string,
  replacementItems: IPermissionAuditItem[]
): IAuditDataCache => ({
  ...cache,
  baseResult: replacePermissionItemInResult(cache.baseResult, itemKey, replacementItems),
  expandedBaseResult: replacePermissionItemInResult(cache.expandedBaseResult, itemKey, replacementItems),
  listItemsLayer: replacePermissionItemInLayer(cache.listItemsLayer, itemKey, replacementItems),
  listPermissionsLayer: replacePermissionItemInLayer(cache.listPermissionsLayer, itemKey, replacementItems)
});

const findDeferredGroupMemberItems = (permissionItems: IPermissionAuditItem[]): IPermissionAuditItem[] =>
  permissionItems.reduce((items: IPermissionAuditItem[], permissionItem) => {
    if (permissionItem.sourceType === 'DeferredGroupMembers' && permissionItem.deferredGroupMembersDetails) {
      items.push(permissionItem);
    }

    if (permissionItem.children?.length) {
      items.push(...findDeferredGroupMemberItems(permissionItem.children));
    }

    return items;
  }, []);

const toUnloadedDeferredGroupMembersItem = (permissionItem: IPermissionAuditItem): IPermissionAuditItem => {
  const path: string[] = permissionItem.path.length
    ? [
      ...permissionItem.path.slice(0, -1),
      strings.AuditRemainingGroupMembersNotLoadedLabel
    ]
    : [strings.AuditRemainingGroupMembersNotLoadedLabel];

  return {
    ...permissionItem,
    key: `${permissionItem.key}|not-loaded`,
    displayName: strings.AuditRemainingGroupMembersNotLoadedLabel,
    principalType: 'Unknown',
    description: strings.AuditRemainingGroupMembersNotLoadedDetails,
    path,
    deferredGroupMembersDetails: undefined,
    children: []
  };
};

const replaceDeferredGroupMemberItems = (permissionItems: IPermissionAuditItem[]): IPermissionAuditItem[] =>
  permissionItems.map((permissionItem) => {
    if (permissionItem.sourceType === 'DeferredGroupMembers' && permissionItem.deferredGroupMembersDetails) {
      return toUnloadedDeferredGroupMembersItem(permissionItem);
    }

    return {
      ...permissionItem,
      children: permissionItem.children?.length
        ? replaceDeferredGroupMemberItems(permissionItem.children)
        : permissionItem.children
    };
  });

const replaceDeferredGroupMemberItemsInResult = (
  auditResult: ICurrentSitePermissionGroupsResult
): ICurrentSitePermissionGroupsResult => ({
  ...auditResult,
  groups: replaceDeferredGroupMemberItems(auditResult.groups)
});

const replaceDeferredGroupMemberItemsInLayer = (
  listLayer: ILoadedListLayer | undefined
): ILoadedListLayer | undefined => listLayer
  ? {
    ...listLayer,
    expandedNodes: listLayer.expandedNodes
      ? replaceDeferredGroupMemberItems(listLayer.expandedNodes)
      : listLayer.expandedNodes,
    nodes: replaceDeferredGroupMemberItems(listLayer.nodes)
  }
  : listLayer;

const replaceDeferredGroupMemberItemsInCache = (cache: IAuditDataCache): IAuditDataCache => ({
  ...cache,
  baseResult: cache.baseResult
    ? replaceDeferredGroupMemberItemsInResult(cache.baseResult)
    : cache.baseResult,
  expandedBaseResult: cache.expandedBaseResult
    ? replaceDeferredGroupMemberItemsInResult(cache.expandedBaseResult)
    : cache.expandedBaseResult,
  listItemsLayer: replaceDeferredGroupMemberItemsInLayer(cache.listItemsLayer),
  listPermissionsLayer: replaceDeferredGroupMemberItemsInLayer(cache.listPermissionsLayer)
});

const loadAllDeferredGroupMembersAsync = async (
  initialAuditResult: ICurrentSitePermissionGroupsResult,
  loadDeferredGroupMembersAsync: (loadMoreItem: IPermissionAuditItem) => Promise<IPermissionAuditItem[]>,
  onDeferredGroupMembersLoaded: (
    itemKey: string,
    loadedItems: IPermissionAuditItem[],
    nextAuditResult: ICurrentSitePermissionGroupsResult
  ) => void
): Promise<ICurrentSitePermissionGroupsResult> => {
  let nextAuditResult: ICurrentSitePermissionGroupsResult = initialAuditResult;
  let nextLoadMoreItem: IPermissionAuditItem | undefined = findDeferredGroupMemberItems(nextAuditResult.groups)[0];

  while (nextLoadMoreItem) {
    const loadedItems: IPermissionAuditItem[] = await loadDeferredGroupMembersAsync(nextLoadMoreItem);
    const replacedAuditResult: ICurrentSitePermissionGroupsResult | undefined =
      replacePermissionItemInResult(nextAuditResult, nextLoadMoreItem.key, loadedItems);

    if (!replacedAuditResult) {
      return nextAuditResult;
    }

    nextAuditResult = replacedAuditResult;
    onDeferredGroupMembersLoaded(nextLoadMoreItem.key, loadedItems, nextAuditResult);
    nextLoadMoreItem = findDeferredGroupMemberItems(nextAuditResult.groups)[0];
  }

  return nextAuditResult;
};

const filterListNodesByHiddenOption = (
  listNodes: IPermissionAuditItem[] | undefined,
  includeHiddenLists: boolean
): IPermissionAuditItem[] => (listNodes || []).filter((listNode) =>
  includeHiddenLists || !listNode.objectDetails?.hidden
);

const getListLayerNodes = (
  listLayer: ILoadedListLayer | undefined,
  includeHiddenLists: boolean,
  includeExpandedGroups: boolean
): IPermissionAuditItem[] => filterListNodesByHiddenOption(
  includeExpandedGroups && listLayer?.expandedNodes ? listLayer.expandedNodes : listLayer?.nodes,
  includeHiddenLists
);

const mergeListLayerNodes = (
  listPermissionNodes: IPermissionAuditItem[],
  listItemNodes: IPermissionAuditItem[]
): IPermissionAuditItem[] => {
  const listNodeByKey: Map<string, IPermissionAuditItem> = new Map<string, IPermissionAuditItem>();

  listPermissionNodes.forEach((listNode) => {
    listNodeByKey.set(listNode.key, {
      ...listNode,
      children: [...(listNode.children || [])]
    });
  });

  listItemNodes.forEach((listNode) => {
    const existingListNode: IPermissionAuditItem | undefined = listNodeByKey.get(listNode.key);

    if (!existingListNode) {
      listNodeByKey.set(listNode.key, {
        ...listNode,
        children: [...(listNode.children || [])]
      });
      return;
    }

    existingListNode.children = [
      ...(existingListNode.children || []),
      ...(listNode.children || []).filter((child) => child.principalType === 'ListItem')
    ];
  });

  return Array.from(listNodeByKey.values());
};

const composeAuditResultFromCache = (
  cache: IAuditDataCache,
  selectedOptionKeys: string[]
): ICurrentSitePermissionGroupsResult | undefined => {
  const includeExpandedGroups: boolean = selectedOptionKeys.indexOf('ExpandGroups') > -1;
  const includeListsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') > -1;
  const includeListItemsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') > -1;
  const includeHiddenLists: boolean = selectedOptionKeys.indexOf('IncludeHiddenListsWithUniquePermissions') > -1;
  const baseResult: ICurrentSitePermissionGroupsResult | undefined = includeExpandedGroups && cache.expandedBaseResult
    ? cache.expandedBaseResult
    : cache.baseResult;
  const siteNode: IPermissionAuditItem | undefined = baseResult?.groups[0];

  if (!baseResult || !siteNode) {
    return undefined;
  }

  const listPermissionNodes: IPermissionAuditItem[] = includeListsWithUniquePermissions
    ? getListLayerNodes(cache.listPermissionsLayer, includeHiddenLists, includeExpandedGroups)
    : [];
  const listItemNodes: IPermissionAuditItem[] = includeListItemsWithUniquePermissions
    ? getListLayerNodes(cache.listItemsLayer, includeHiddenLists, includeExpandedGroups)
    : [];
  const mergedListNodes: IPermissionAuditItem[] = mergeListLayerNodes(listPermissionNodes, listItemNodes);

  return {
    associatedGroupIds: baseResult.associatedGroupIds,
    groups: [{
      ...siteNode,
      children: [
        ...(siteNode.children || []),
        ...mergedListNodes
      ]
    }]
  };
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

const toPrincipalSuggestionPersona = (directoryPerson: IDirectoryPersonInfo): IPrincipalSuggestionPersona => ({
  key: directoryPerson.id,
  text: directoryPerson.displayName,
  primaryText: directoryPerson.displayName,
  secondaryText: directoryPerson.userPrincipalName || directoryPerson.email,
  tertiaryText: directoryPerson.department,
  optionalText: directoryPerson.jobTitle,
  directoryPerson
});

export const AuditPage: React.FunctionComponent<IAuditPageProps> = (props) => {
  const [selectedOptionKeys, setSelectedOptionKeys] = React.useState<string[]>([]);
  const [auditGroups, setAuditGroups] = React.useState<IGroup[] | undefined>(undefined);
  const [auditItems, setAuditItems] = React.useState<IAuditGridItem[]>([]);
  const [auditResult, setAuditResult] = React.useState<ICurrentSitePermissionGroupsResult | undefined>(undefined);
  const [auditCache, setAuditCache] = React.useState<IAuditDataCache>({});
  const [selectedDetailsHtml, setSelectedDetailsHtml] = React.useState<string>('');
  const [selectedPrincipalSuggestions, setSelectedPrincipalSuggestions] = React.useState<IPrincipalSuggestionPersona[]>([]);
  const [isAuditLoading, setIsAuditLoading] = React.useState<boolean>(false);
  const [isPrincipalSearchLoading, setIsPrincipalSearchLoading] = React.useState<boolean>(false);
  const [isGroupedView, setIsGroupedView] = React.useState<boolean>(() => readGroupedViewPreference(props.groupedViewPreferenceKey));
  const loadDeferredGroupMembersRef = React.useRef<(item: IPermissionAuditItem) => void>(() => undefined);
  const detailsPanelHeightPreferenceKey: string = React.useMemo(
    () => `${props.groupedViewPreferenceKey}:GridHeight`,
    [props.groupedViewPreferenceKey]
  );
  const [auditGridHeight, setAuditGridHeight] = React.useState<number>(() => readPanelHeightPreference(`${props.groupedViewPreferenceKey}:GridHeight`));
  const isSharePointReady: boolean = !!props.sharePointPermissionAuditService;
  const selectionRef = React.useRef<Selection>();
  const latestAuditRequestIdRef = React.useRef<number>(0);
  const latestPrincipalSearchRequestIdRef = React.useRef<number>(0);
  const resizeStateRef = React.useRef<{ startY: number; startHeight: number } | undefined>(undefined);
  const includeListsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') > -1;
  const includeListItemsWithUniquePermissions: boolean = selectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') > -1;
  const includeListAuditScope: boolean = includeListsWithUniquePermissions || includeListItemsWithUniquePermissions;
  const maximumGridHeight: number = auditSplitPaneHeight - auditResizeHandleHeight - auditMinimumDetailsHeight;
  const auditDetailsContentHeight: number = auditSplitPaneHeight - auditResizeHandleHeight - auditGridHeight;
  const isLoading: boolean = isAuditLoading || isPrincipalSearchLoading;
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
        const selectedItem: IAuditGridItem | undefined = selectedItems[0];

        if (selectedItem?.permissionItem.principalType === 'LoadMore') {
          loadDeferredGroupMembersRef.current(selectedItem.permissionItem);
          return;
        }

        setSelectedDetailsHtml(selectedItem?.detailsHtml || '');
      }
    });
  }

  const applyAuditCacheProjection = React.useCallback((nextCache: IAuditDataCache, nextSelectedOptionKeys: string[]): void => {
    const nextAuditResult: ICurrentSitePermissionGroupsResult | undefined =
      composeAuditResultFromCache(nextCache, nextSelectedOptionKeys);

    if (!nextAuditResult) {
      setAuditResult(undefined);
      setAuditGroups(undefined);
      setAuditItems([]);
      return;
    }

    const projection: IAuditGridProjection = projectAuditResult(nextAuditResult, isGroupedView);

    setAuditResult(nextAuditResult);
    setAuditGroups(projection.groups);
    setAuditItems(projection.items);
  }, [isGroupedView]);

  const loadMissingAuditLayers = React.useCallback(async (requestId: number, nextSelectedOptionKeys: string[]): Promise<void> => {
    const shouldExpandGroups: boolean = nextSelectedOptionKeys.indexOf('ExpandGroups') > -1;
    const shouldIncludeListsWithUniquePermissions: boolean = nextSelectedOptionKeys.indexOf('IncludeListsWithUniquePermissions') > -1;
    const shouldIncludeHiddenListsWithUniquePermissions: boolean = nextSelectedOptionKeys.indexOf('IncludeHiddenListsWithUniquePermissions') > -1;
    const shouldIncludeListItemsWithUniquePermissions: boolean = nextSelectedOptionKeys.indexOf('IncludeListItemsWithUniquePermissions') > -1;

    setIsAuditLoading(true);

    try {
      let nextCache: IAuditDataCache = auditCache;

      if (!nextCache.baseResult) {
        const baseResult: ICurrentSitePermissionGroupsResult =
          await props.sharePointPermissionAuditService.getPermissionAuditAsync({
            expandGroups: false,
            groupExpansionBatchSize: props.groupExpansionBatchSize
          });

        nextCache = {
          ...nextCache,
          baseResult
        };
      }

      if (shouldExpandGroups && !nextCache.expandedBaseResult) {
        const siteNode: IPermissionAuditItem | undefined = nextCache.baseResult?.groups[0];
        const expandedChildren: IPermissionAuditItem[] =
          await props.sharePointPermissionAuditService.expandPermissionAuditGroupsAsync(
            siteNode?.children || [],
            props.groupExpansionBatchSize
          );

        nextCache = {
          ...nextCache,
          expandedBaseResult: nextCache.baseResult && siteNode ? {
            associatedGroupIds: nextCache.baseResult.associatedGroupIds,
            groups: [{
              ...siteNode,
              children: expandedChildren
            }]
          } : undefined
        };
      }

      if (
        shouldIncludeListsWithUniquePermissions &&
        (!nextCache.listPermissionsLayer || (shouldIncludeHiddenListsWithUniquePermissions && !nextCache.listPermissionsLayer.includeHidden))
      ) {
        const listPermissionNodes: IPermissionAuditItem[] =
          await props.sharePointPermissionAuditService.getListPermissionAuditItemsAsync({
            expandGroups: false,
            groupExpansionBatchSize: props.groupExpansionBatchSize,
            includeHiddenListsWithUniquePermissions: shouldIncludeHiddenListsWithUniquePermissions,
            includeListsWithUniquePermissions: true
          });

        nextCache = {
          ...nextCache,
          listPermissionsLayer: {
            includeHidden: shouldIncludeHiddenListsWithUniquePermissions,
            nodes: listPermissionNodes
          }
        };
      }

      if (
        shouldExpandGroups &&
        nextCache.listPermissionsLayer &&
        !nextCache.listPermissionsLayer.expandedNodes
      ) {
        const expandedListPermissionNodes: IPermissionAuditItem[] =
          await props.sharePointPermissionAuditService.expandPermissionAuditGroupsAsync(
            nextCache.listPermissionsLayer.nodes,
            props.groupExpansionBatchSize
          );

        nextCache = {
          ...nextCache,
          listPermissionsLayer: {
            ...nextCache.listPermissionsLayer,
            expandedNodes: expandedListPermissionNodes
          }
        };
      }

      if (
        shouldIncludeListItemsWithUniquePermissions &&
        (!nextCache.listItemsLayer || (shouldIncludeHiddenListsWithUniquePermissions && !nextCache.listItemsLayer.includeHidden))
      ) {
        const listItemNodes: IPermissionAuditItem[] =
          await props.sharePointPermissionAuditService.getListPermissionAuditItemsAsync({
            expandGroups: false,
            groupExpansionBatchSize: props.groupExpansionBatchSize,
            includeHiddenListsWithUniquePermissions: shouldIncludeHiddenListsWithUniquePermissions,
            includeListItemsWithUniquePermissions: true
          });

        nextCache = {
          ...nextCache,
          listItemsLayer: {
            includeHidden: shouldIncludeHiddenListsWithUniquePermissions,
            nodes: listItemNodes
          }
        };
      }

      if (
        shouldExpandGroups &&
        nextCache.listItemsLayer &&
        !nextCache.listItemsLayer.expandedNodes
      ) {
        const expandedListItemNodes: IPermissionAuditItem[] =
          await props.sharePointPermissionAuditService.expandPermissionAuditGroupsAsync(
            nextCache.listItemsLayer.nodes,
            props.groupExpansionBatchSize
          );

        nextCache = {
          ...nextCache,
          listItemsLayer: {
            ...nextCache.listItemsLayer,
            expandedNodes: expandedListItemNodes
          }
        };
      }

      if (latestAuditRequestIdRef.current !== requestId) {
        return;
      }

      setAuditCache(nextCache);
      applyAuditCacheProjection(nextCache, nextSelectedOptionKeys);
    } catch (error) {
      if (latestAuditRequestIdRef.current === requestId) {
        setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
      }
    } finally {
      if (latestAuditRequestIdRef.current === requestId) {
        setIsAuditLoading(false);
      }
    }
  }, [
    applyAuditCacheProjection,
    auditCache,
    props.groupExpansionBatchSize,
    props.sharePointPermissionAuditService
  ]);

  React.useEffect(() => {
    if (!isSharePointReady) {
      return undefined;
    }

    const requestId: number = latestAuditRequestIdRef.current + 1;
    latestAuditRequestIdRef.current = requestId;

    const timeoutHandle: number = window.setTimeout(() => {
      loadMissingAuditLayers(requestId, selectedOptionKeys).catch((error) => {
        if (latestAuditRequestIdRef.current === requestId) {
          setIsAuditLoading(false);
          setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
        }
      });
    }, auditAutoRunDebounceMs);

    return () => {
      window.clearTimeout(timeoutHandle);
    };
  }, [isSharePointReady, loadMissingAuditLayers, selectedOptionKeys]);

  React.useEffect(() => {
    setAuditGridHeight(readPanelHeightPreference(detailsPanelHeightPreferenceKey));
  }, [detailsPanelHeightPreferenceKey]);

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

  const onExportCsvClick = React.useCallback(async (): Promise<void> => {
    if (!auditResult) {
      return;
    }

    const deferredGroupMemberItems: IPermissionAuditItem[] = findDeferredGroupMemberItems(auditResult.groups);

    if (!deferredGroupMemberItems.length) {
      downloadCsv(buildCsvContent(auditResult));
      return;
    }

    const shouldLoadRemainingMembers: boolean = window.confirm(strings.AuditExportLoadRemainingGroupMembersPrompt);

    if (!shouldLoadRemainingMembers) {
      const nextAuditResult: ICurrentSitePermissionGroupsResult = replaceDeferredGroupMemberItemsInResult(auditResult);
      const projection: IAuditGridProjection = projectAuditResult(nextAuditResult, isGroupedView);

      setAuditCache((currentCache) => replaceDeferredGroupMemberItemsInCache(currentCache));
      setAuditResult(nextAuditResult);
      setAuditGroups(projection.groups);
      setAuditItems(projection.items);
      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditRemainingGroupMembersNotLoadedDetails)}</p>`);
      downloadCsv(buildCsvContent(nextAuditResult));
      return;
    }

    setIsAuditLoading(true);
    setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditExportLoadingRemainingGroupMembersLabel)}</p>`);

    try {
      const nextAuditResult: ICurrentSitePermissionGroupsResult = await loadAllDeferredGroupMembersAsync(
        auditResult,
        (loadMoreItem) => props.sharePointPermissionAuditService.loadDeferredGroupMembersAsync(loadMoreItem),
        (itemKey, loadedItems, updatedAuditResult) => {
          const projection: IAuditGridProjection = projectAuditResult(updatedAuditResult, isGroupedView);

          setAuditCache((currentCache) => replacePermissionItemInCache(currentCache, itemKey, loadedItems));
          setAuditResult(updatedAuditResult);
          setAuditGroups(projection.groups);
          setAuditItems(projection.items);
        }
      );

      downloadCsv(buildCsvContent(nextAuditResult));
      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditLoadRemainingGroupMembersDetails)}</p>`);
    } catch (error) {
      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
    } finally {
      setIsAuditLoading(false);
    }
  }, [auditResult, isGroupedView, props.sharePointPermissionAuditService]);

  const onLoadDeferredGroupMembers = React.useCallback((loadMoreItem: IPermissionAuditItem): void => {
    if (!auditResult) {
      return;
    }

    setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditLoadingLabel)}</p>`);

    props.sharePointPermissionAuditService.loadDeferredGroupMembersAsync(loadMoreItem)
      .then((loadedItems) => {
        const nextAuditResult: ICurrentSitePermissionGroupsResult | undefined =
          replacePermissionItemInResult(auditResult, loadMoreItem.key, loadedItems);

        if (!nextAuditResult) {
          return;
        }

        const projection: IAuditGridProjection = projectAuditResult(nextAuditResult, isGroupedView);

        setAuditCache((currentCache) => replacePermissionItemInCache(currentCache, loadMoreItem.key, loadedItems));
        setAuditResult(nextAuditResult);
        setAuditGroups(projection.groups);
        setAuditItems(projection.items);
        setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditLoadRemainingGroupMembersDetails)}</p>`);
      })
      .catch((error) => {
        setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
      });
  }, [auditResult, isGroupedView, props.sharePointPermissionAuditService]);

  loadDeferredGroupMembersRef.current = onLoadDeferredGroupMembers;

  const onResolvePrincipalSuggestions = React.useCallback(async (
    filterText: string,
    currentSuggestions?: IPersonaProps[]
  ): Promise<IPrincipalSuggestionPersona[]> => {
    if (filterText.trim().length < 3) {
      return [];
    }

    const directoryPeople: IDirectoryPersonInfo[] = await props.graphPermissionAuditService.searchPeopleAsync(filterText);
    const selectedSuggestionKeys: Set<string> = new Set((currentSuggestions || []).map((suggestion) => `${suggestion.key}`));

    return directoryPeople
      .filter((directoryPerson) => !selectedSuggestionKeys.has(directoryPerson.id))
      .map(toPrincipalSuggestionPersona);
  }, [props.graphPermissionAuditService]);

  const onPrincipalSelectionChange = React.useCallback((items?: IPersonaProps[]): void => {
    setSelectedPrincipalSuggestions(((items as IPrincipalSuggestionPersona[] | undefined) || []).slice(0, 1));
  }, []);

  const onPrincipalSearchClear = React.useCallback((): void => {
    latestPrincipalSearchRequestIdRef.current += 1;
    setSelectedPrincipalSuggestions([]);
    setSelectedDetailsHtml('');
  }, []);

  React.useEffect(() => {
    const selectedPrincipal: IPrincipalSuggestionPersona | undefined = selectedPrincipalSuggestions[0];

    if (!selectedPrincipal) {
      return;
    }

    const requestId: number = latestPrincipalSearchRequestIdRef.current + 1;
    latestPrincipalSearchRequestIdRef.current = requestId;
    setIsPrincipalSearchLoading(true);
    setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditLoadingLabel)}</p>`);

    props.sharePointPermissionAuditService.searchPrincipalAccessAsync({
      groupExpansionBatchSize: props.groupExpansionBatchSize,
      principal: selectedPrincipal.directoryPerson,
      includeHiddenListsWithUniquePermissions: selectedOptionKeys.indexOf('IncludeHiddenListsWithUniquePermissions') > -1
    }).then((searchResult) => {
      if (latestPrincipalSearchRequestIdRef.current !== requestId) {
        return;
      }

      setSelectedDetailsHtml(buildPrincipalSearchDetailsHtml(
        searchResult,
        searchResult.auditResult.associatedGroupIds
      ));
    }).catch((error) => {
      if (latestPrincipalSearchRequestIdRef.current !== requestId) {
        return;
      }

      setSelectedDetailsHtml(`<p>${encodeHtml(strings.AuditGroupFetchError)}</p><p>${encodeHtml(error instanceof Error ? error.message : '')}</p>`);
    }).then(() => {
      if (latestPrincipalSearchRequestIdRef.current === requestId) {
        setIsPrincipalSearchLoading(false);
      }
    }, () => {
      if (latestPrincipalSearchRequestIdRef.current === requestId) {
        setIsPrincipalSearchLoading(false);
      }
    });
  }, [props.groupExpansionBatchSize, props.sharePointPermissionAuditService, selectedOptionKeys, selectedPrincipalSuggestions]);

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

  const updateGridHeight = React.useCallback((nextGridHeight: number): void => {
    const normalizedHeight: number = clamp(nextGridHeight, auditMinimumGridHeight, maximumGridHeight);

    setAuditGridHeight(normalizedHeight);
    writePanelHeightPreference(detailsPanelHeightPreferenceKey, normalizedHeight);
  }, [detailsPanelHeightPreferenceKey, maximumGridHeight]);

  const onResizeHandlePointerDown = React.useCallback((event: React.PointerEvent<HTMLButtonElement>): void => {
    resizeStateRef.current = {
      startY: event.clientY,
      startHeight: auditGridHeight
    };
    event.currentTarget.setPointerCapture(event.pointerId);
    event.preventDefault();
  }, [auditGridHeight]);

  const onResizeHandlePointerMove = React.useCallback((event: React.PointerEvent<HTMLButtonElement>): void => {
    if (!resizeStateRef.current) {
      return;
    }

    updateGridHeight(resizeStateRef.current.startHeight + (event.clientY - resizeStateRef.current.startY));
  }, [updateGridHeight]);

  const onResizeHandlePointerUp = React.useCallback((event: React.PointerEvent<HTMLButtonElement>): void => {
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }

    resizeStateRef.current = undefined;
  }, []);

  const onResizeHandlePointerCancel = React.useCallback((event: React.PointerEvent<HTMLButtonElement>): void => {
    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId);
    }

    resizeStateRef.current = undefined;
  }, []);

  const onResizeHandleKeyDown = React.useCallback((event: React.KeyboardEvent<HTMLButtonElement>): void => {
    const resizeStep: number = event.shiftKey ? 48 : 24;

    switch (event.key) {
      case 'ArrowUp':
        updateGridHeight(auditGridHeight - resizeStep);
        event.preventDefault();
        return;
      case 'ArrowDown':
        updateGridHeight(auditGridHeight + resizeStep);
        event.preventDefault();
        return;
      case 'Home':
        updateGridHeight(auditMinimumGridHeight);
        event.preventDefault();
        return;
      case 'End':
        updateGridHeight(maximumGridHeight);
        event.preventDefault();
        return;
      default:
        return;
    }
  }, [auditGridHeight, maximumGridHeight, updateGridHeight]);

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
        <NormalPeoplePicker
          aria-label={strings.AuditPrincipalSearchLabel}
          disabled={isLoading}
          inputProps={{
            'aria-label': strings.AuditPrincipalSearchLabel,
            placeholder: strings.AuditPrincipalSearchPlaceholder
          }}
          itemLimit={1}
          onChange={onPrincipalSelectionChange}
          onResolveSuggestions={onResolvePrincipalSuggestions}
          pickerSuggestionsProps={{
            noResultsFoundText: strings.AuditPrincipalSearchNoResults
          }}
          resolveDelay={400}
          selectedItems={selectedPrincipalSuggestions}
        />
        <DefaultButton
          text={strings.AuditPrincipalSearchClearButtonLabel}
          disabled={selectedPrincipalSuggestions.length === 0}
          onClick={onPrincipalSearchClear}
        />
      </Stack>

      <div className={styles.splitPane} style={{ height: auditSplitPaneHeight }}>
        <div className={styles.gridPane} style={{ height: auditGridHeight }}>
          <div className={styles.grid}>
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
            {isLoading && auditItems.length === 0 && (
              <div className={styles.gridLoadingOverlay}>
                <Spinner
                  label={strings.AuditLoadingLabel}
                  size={SpinnerSize.large}
                />
              </div>
            )}
          </div>
        </div>

        <button
          aria-label={strings.AuditSplitResizeHandleAriaLabel}
          aria-orientation="horizontal"
          aria-valuemax={maximumGridHeight}
          aria-valuemin={auditMinimumGridHeight}
          aria-valuenow={auditGridHeight}
          className={styles.resizeHandle}
          onKeyDown={onResizeHandleKeyDown}
          onPointerCancel={onResizeHandlePointerCancel}
          onPointerDown={onResizeHandlePointerDown}
          onPointerMove={onResizeHandlePointerMove}
          onPointerUp={onResizeHandlePointerUp}
          role="separator"
          type="button"
        />

        <div className={styles.detailsPane} style={{ height: auditDetailsContentHeight }}>
          <HtmlContentDisplay
            ariaLabel={strings.AuditDetailsContentAriaLabel}
            className={styles.detailsContent}
            height={auditDetailsContentHeight}
            html={selectedDetailsHtml}
          />
        </div>
      </div>
    </section>
  );
};
