import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { loadTheme } from '@fluentui/react/lib/Styling';
import { MSGraphClientFactory } from '@microsoft/sp-http-msgraph';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';

import * as strings from 'PermissionAuditWebPartStrings';
import PermissionAudit from './components/PermissionAudit';
import { IPermissionAuditProps } from './components/IPermissionAuditProps';
import {
  GraphPermissionAuditService,
  IGraphPermissionAuditService,
  ISharePointPermissionAuditService,
  SharePointPermissionAuditService
} from './services';

export interface IPermissionAuditWebPartProps {
  description: string;
  groupExpansionBatchSize: number;
}

export default class PermissionAuditWebPart extends BaseClientSideWebPart<IPermissionAuditWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp!: SPFI;
  private _graphPermissionAuditService!: IGraphPermissionAuditService;
  private _sharePointPermissionAuditService!: ISharePointPermissionAuditService;

  public render(): void {
    const element: React.ReactElement<IPermissionAuditProps> = React.createElement(
      PermissionAudit,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        graphPermissionAuditService: this._graphPermissionAuditService,
        groupExpansionBatchSize: this._getGroupExpansionBatchSize(),
        groupedViewPreferenceKey: this._getGroupedViewPreferenceKey(),
        sharePointPermissionAuditService: this._sharePointPermissionAuditService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    this._graphPermissionAuditService = new GraphPermissionAuditService(
      this.context.serviceScope.consume(MSGraphClientFactory.serviceKey)
    );
    this._sharePointPermissionAuditService = new SharePointPermissionAuditService(this._sp, this._graphPermissionAuditService);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  private _getGroupedViewPreferenceKey(): string {
    const tenantId: string = this.context.pageContext.aadInfo?.tenantId?._guid || this.context.pageContext.site.absoluteUrl;
    const userLoginName: string = this.context.pageContext.user.loginName || this.context.pageContext.user.email || 'current-user';

    return `PermissionAudit:${tenantId}:${userLoginName}:GroupedView`;
  }

  private _getGroupExpansionBatchSize(): number {
    const configuredBatchSize: number = Number(this.properties.groupExpansionBatchSize);

    if (!configuredBatchSize || configuredBatchSize < 1) {
      return 100;
    }

    return Math.min(Math.floor(configuredBatchSize), 5000);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    loadTheme(currentTheme);

    const {
      palette,
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--bodySubtext', semanticColors.bodySubtext || null);
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--bodyStandoutBackground', semanticColors.bodyStandoutBackground || null);
      this.domElement.style.setProperty('--bodyDivider', semanticColors.bodyDivider || null);
      this.domElement.style.setProperty('--inputBorder', semanticColors.inputBorder || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
      this.domElement.style.setProperty('--neutralLighter', palette.neutralLighter || null);
      this.domElement.style.setProperty('--neutralQuaternaryAlt', palette.neutralQuaternaryAlt || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('groupExpansionBatchSize', {
                  label: strings.GroupExpansionBatchSizeFieldLabel,
                  description: strings.GroupExpansionBatchSizeFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
