import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  PropertyFieldFilePicker,
  IFilePickerResult
} from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldColorPicker,
  PropertyFieldColorPickerStyle
} from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import * as strings from 'GreetingsWebPartStrings';
import Greetings from './components/Greetings';
import { IGreetingsProps } from './components/IGreetingsProps';

export interface IGreetingsWebPartProps {
  description: string;
  greetingText: string;
  showFirstNameOnly: boolean;
  fontSize: string;
  fontSizeUnit: string;
  fontSizeValue: string;
  fontStyle: string;
  backgroundImageUrl: string;
  fontColor?: string; // ✅ Make optional
  textAlignment: string; // ✅ ADD THIS LINE
}

export default class GreetingsWebPart extends BaseClientSideWebPart<IGreetingsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // Font size units options
  private readonly fontSizeUnits = [
    { key: 'px', text: 'px (pixels)' },
    { key: 'pt', text: 'pt (points)' },
  ];

  // Font size predefined values
  private readonly fontSizeValues = [
    { key: '12', text: '12' },
    { key: '14', text: '14' },
    { key: '16', text: '16' },
    { key: '18', text: '18' },
    { key: '20', text: '20' },
    { key: '24', text: '24' },
    { key: '28', text: '28' },
    { key: '32', text: '32' },
    { key: '36', text: '36' },
    { key: '40', text: '40' },
    { key: '48', text: '48' },
    { key: 'custom', text: 'Custom' }
  ];

  public render(): void {
    // Combine font size value and unit
    const fontSize = this.getFontSize();

    const element: React.ReactElement<IGreetingsProps> = React.createElement(
      Greetings,
      {
        description: this.properties.description,
        greetingText: this.properties.greetingText || 'Welcome',
        showFirstNameOnly: this.properties.showFirstNameOnly,
        fontSize: fontSize,
        fontStyle: this.properties.fontStyle || 'Normal',
        backgroundImageUrl: this.properties.backgroundImageUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context:this.context,
         textAlignment: this.properties.textAlignment || 'center', // ✅ NEW: Pass alignment (default center)
      fontColor: this.properties.fontColor || '', // ✅ Empty string if no color selected

      }
    );

    ReactDom.render(element, this.domElement);
  }

  // Helper method to combine font size
  private getFontSize(): string {
    const { fontSizeValue, fontSizeUnit } = this.properties;

    if (!fontSizeValue || fontSizeValue === 'custom') {
      return '12px';
    }

    return `${fontSizeValue}${fontSizeUnit || 'px'}`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
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

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private onFilePickerSave = (filePickerResult: IFilePickerResult): string => {
    if (filePickerResult) {
      let imageUrl = '';

      // Debug log
      console.log('FilePicker Result:', filePickerResult);

      if (filePickerResult.fileAbsoluteUrl) {
        // Use the absolute URL if available
        imageUrl = filePickerResult.fileAbsoluteUrl;
      } else if (filePickerResult.spItemUrl) {
        // For SharePoint items
        imageUrl = filePickerResult.spItemUrl;
      } else if (filePickerResult.fileName) {
        // Fallback to the first file
        imageUrl = filePickerResult.fileName;
      } else if (typeof filePickerResult === 'string') {
        // If it's already a string
        imageUrl = filePickerResult;
      }

      console.log('Selected image URL:', imageUrl);
      return imageUrl;
    }

    return '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
      
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('greetingText', {
                  label: 'Greeting Text',
                  value: 'Welcome',
                  description: 'Enter greeting text (e.g., Welcome, Hello, Hi)'
                }),
                PropertyPaneToggle('showFirstNameOnly', {
                  label: 'Show First Name Only',
                  onText: 'First Name',
                  offText: 'Full Name',
                  checked: false
                }),
                // Font Size Section
                PropertyPaneDropdown('fontSizeValue', {
                  label: 'Font Size',
                  selectedKey: '16',
                  options: this.fontSizeValues
                }),
                   PropertyFieldColorPicker('fontColor', {
                label: 'Font Color',
                selectedColor: this.properties.fontColor || '',
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.properties,
                disabled: false,
                isHidden: false,
                alphaSliderHidden: true,
                style: PropertyFieldColorPickerStyle.Inline,
                iconName: 'Color',
                key: 'colorPickerFieldId'
              }),
              // ✅ NEW: Text Alignment
              PropertyPaneDropdown('textAlignment', {
                label: 'Text Alignment',
                selectedKey: this.properties.textAlignment || 'center',
                options: [
                  { key: 'left', text: 'Left' },
                  { key: 'center', text: 'Center' },
                  { key: 'right', text: 'Right' }
                ]
              }),
                PropertyPaneDropdown('fontSizeUnit', {
                  label: 'Font Unit',
                  selectedKey: 'px',
                  options: this.fontSizeUnits
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: 'Font Style',
                  selectedKey: 'Normal',
                  options: [
                    { key: 'Normal', text: 'Normal' },
                    { key: 'Bold', text: 'Bold' },
                    { key: 'Italic', text: 'Italic' },
                    { key: 'Bold Italic', text: 'Bold Italic' }
                  ]
                }),
                PropertyFieldFilePicker('backgroundImageUrl', {
                  context: this.context,

                  // Fix: Handle both string and object cases
                  filePickerResult: this.properties.backgroundImageUrl
                    ? (typeof this.properties.backgroundImageUrl === 'string'
                      ? ({
                        fileAbsoluteUrl: this.properties.backgroundImageUrl,
                        fileName: this.properties.backgroundImageUrl.split('/').pop() || ''
                      } as IFilePickerResult)
                      : this.properties.backgroundImageUrl)
                    : ({} as IFilePickerResult),

                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'backgroundImagePicker',

                  buttonLabel: 'Browse Image',
                  label: 'Background Image',

                  accepts: ['.png', '.jpg', '.jpeg', '.svg', '.gif'],

                  hideLinkUploadTab: false,
                  hideLocalUploadTab: true,
                  hideWebSearchTab: true,
                  hideStockImages: true,
                  hideRecentTab: true,
                  hideSiteFilesTab: false,
                  allowExternalLinks: true,

                  onSave: (filePickerResult: IFilePickerResult) => {
                    if (filePickerResult) {
                      const imageUrl = this.onFilePickerSave(filePickerResult);
                      if (imageUrl) {
                        this.properties.backgroundImageUrl = imageUrl;
                        this.onPropertyPaneFieldChanged('backgroundImageUrl', '', imageUrl);
                        this.render();
                      }
                    }
                  },

                  onChanged: (filePickerResult: IFilePickerResult) => {
                    if (filePickerResult) {
                      const imageUrl = this.onFilePickerSave(filePickerResult);
                      if (imageUrl) {
                        this.properties.backgroundImageUrl = imageUrl;
                        this.onPropertyPaneFieldChanged('backgroundImageUrl', '', imageUrl);
                        this.render();
                      }
                    }
                  }
                })

              ]
            }
          ]
        }
      ]
    };
  }
}