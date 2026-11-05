import { sp } from "@pnp/sp/presets/all";
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { ISPFXContext } from "@pnp/common";

import * as strings from 'FormularioArtesWebPartStrings';
import FormularioArtes from './components/FormularioArtes';
import { IFormularioArtesProps } from './components/IFormularioArtesProps';

// 👇 IMPORTANTE
import { FormProvider } from './components/approval-form/context/FormContext';

export interface IFormularioArtesWebPartProps {
  description: string;
}

export default class FormularioArtesWebPart
  extends BaseClientSideWebPart<IFormularioArtesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  /* =========================================================
     Inicialización PnPjs (una sola vez)
     ========================================================= */
  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context as unknown as ISPFXContext
      });

      return;
    });
  }

  /* =========================================================
     Render principal
     ========================================================= */
  public render(): void {
    const formProps: IFormularioArtesProps = {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks?.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      spfxContext: this.context
    };

    const element: React.ReactElement = React.createElement(
      FormProvider,
      { spfxContext: this.context },
      React.createElement(FormularioArtes, formProps)
    );

    ReactDom.render(element, this.domElement);
  }

  /* =========================================================
     Tema
     ========================================================= */
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /* =========================================================
     Property Pane
     ========================================================= */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
