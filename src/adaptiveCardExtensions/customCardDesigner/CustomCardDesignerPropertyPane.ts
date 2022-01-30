import { AdaptiveCardExtensionContext, BaseAdaptiveCardExtension, CardSize } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneField,
  IPropertyPaneGroup,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';
import { CustomCollectionFieldType, PropertyFieldCollectionData } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { IFilePickerResult, PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { PropertyPanePropertyEditor } from '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor';

import { CustomCardDesignerAdaptiveCardExtensionProps } from './CustomCardDesignerAdaptiveCardExtension';
import { CardIconSourceType, CardImageSourceType, CardTemplateType } from './CustomCardDesignerTypes';

import * as strings from 'CustomCardDesignerAdaptiveCardExtensionStrings';

const HeadingIcon: string = require('./assets/HeadingIcon.svg');
const ImageIcon: string = require('./assets/ImageIcon.svg');
const DescriptionIcon: string = require('./assets/DescriptionIcon.svg');

export class CustomCardDesignerPropertyPane {
  public getPropertyPaneConfiguration({
    adaptiveCardExtension,
    cardSize,
    context,
    onPropertyPaneFieldChanged,
    properties
  }: {
    adaptiveCardExtension: BaseAdaptiveCardExtension;
    cardSize: CardSize;
    context: AdaptiveCardExtensionContext;
    onPropertyPaneFieldChanged: BaseAdaptiveCardExtension['onPropertyPaneFieldChanged'];
    properties: CustomCardDesignerAdaptiveCardExtensionProps;
  }): IPropertyPaneConfiguration {
    //#region Size and layout group
    const sizeAndLayoutGroup: IPropertyPaneGroup = {
      groupName: strings.LayoutGroupName,
      groupFields: []
    };

    sizeAndLayoutGroup.groupFields.push(
      PropertyPaneChoiceGroup('templateType', {
        label: strings.TemplateTypePropertyLabel,
        options: [
          {
            key: CardTemplateType.Heading,
            text: strings.HeadingTemplateTypeLabel,
            imageSrc: HeadingIcon,
            selectedImageSrc: HeadingIcon
          },
          {
            key: CardTemplateType.Image,
            text: strings.ImageTemplateTypeLabel,
            imageSrc: ImageIcon,
            selectedImageSrc: ImageIcon
          },
          {
            key: CardTemplateType.Description,
            text: strings.DescriptionTemplateTypeLabel,
            imageSrc: DescriptionIcon,
            selectedImageSrc: DescriptionIcon
          }
        ]
      })
    );
    //#endregion

    //#region Card content group
    const cardContentGroup: IPropertyPaneGroup = {
      groupName: strings.CardContentGroupName,
      groupFields: []
    };

    cardContentGroup.groupFields.push(
      PropertyPaneTextField('title', {
        label: strings.CardContentTitlePropertyLabel
      })
    );

    cardContentGroup.groupFields.push(
      PropertyPaneChoiceGroup('cardIconSourceType', {
        label: strings.CardContentIconPropertyLabel,
        options: [
          { key: CardIconSourceType.CustomImage, text: strings.CardContentIconCustomImageLabel },
          { key: CardIconSourceType.Icon, text: strings.CardContentIconIconLabel }
        ]
      })
    );

    if (properties.cardIconSourceType === CardIconSourceType.CustomImage) {
      cardContentGroup.groupFields.push(
        PropertyFieldFilePicker('cardIconFilePickerResult', {
          context,
          filePickerResult: properties.cardIconFilePickerResult,
          onPropertyChange: onPropertyPaneFieldChanged,
          properties: properties,
          onSave: (e: IFilePickerResult) => {
            properties.iconProperty = e.fileAbsoluteUrl;
          },
          onChanged: (e: IFilePickerResult) => {
            properties.iconProperty = e.fileAbsoluteUrl;
          },
          hideWebSearchTab: true,
          hideSiteFilesTab: true,
          accepts: ['.gif', '.jpg', '.jpeg', '.png'],
          key: 'PropertyFieldFilePickerCardIcon',
          buttonLabel: strings.CardContentIconChangeButton
        })
      );
    }

    if (properties.cardIconSourceType === CardIconSourceType.Icon) {
      // @todo: Replace with icon picker
      cardContentGroup.groupFields.push(
        PropertyPaneTextField('iconPicker', {
          label: strings.CardContentIconIconNameLabel
        })
      );
    }

    cardContentGroup.groupFields.push(
      PropertyPaneTextField('primaryText', {
        label: strings.CardContentHeadingPropertyLabel
      })
    );

    if (properties.templateType === CardTemplateType.Image) {
      cardContentGroup.groupFields.push(
        PropertyPaneChoiceGroup('cardImageSourceType', {
          label: strings.CardContentImagePropertyLabel,
          options: [{ key: CardImageSourceType.CustomImage, text: strings.CardContentImageCustomImageLabel }]
        })
      );

      cardContentGroup.groupFields.push(
        PropertyFieldFilePicker('cardImageFilePickerResult', {
          context,
          filePickerResult: properties.cardImageFilePickerResult,
          onPropertyChange: onPropertyPaneFieldChanged,
          properties: properties,
          onSave: (e: IFilePickerResult) => {
            properties.imagePicker = e.fileAbsoluteUrl;
          },
          onChanged: (e: IFilePickerResult) => {
            properties.imagePicker = e.fileAbsoluteUrl;
          },
          hideWebSearchTab: true,
          hideSiteFilesTab: true,
          accepts: ['.gif', '.jpg', '.jpeg', '.png'],
          key: 'PropertyFieldFilePickerCardImage',
          buttonLabel: strings.CardContentImageChangeButton
        })
      );
    }

    if (properties.templateType === CardTemplateType.Description) {
      cardContentGroup.groupFields.push(
        PropertyPaneTextField('description', {
          label: strings.CardContentDescriptionPropertyLabel
        })
      );
    }
    //#endregion

    //#region Card actions group
    const actionsGroup: IPropertyPaneGroup = {
      groupName: strings.ActionsGroupName,
      groupFields: []
    };

    actionsGroup.groupFields.push(
      PropertyPaneDropdown('cardSelectionAction.type', {
        label: strings.CardActionPropertyLabel,
        options: [
          {
            key: 'QuickView',
            text: strings.ShowTheQuickViewLabel
          },
          {
            key: 'ExternalLink',
            text: strings.GoToALinkLabel
          },
          {
            key: 'TeamsExternalLink',
            text: strings.GoToATeamsAppLabel
          }
        ]
      })
    );

    if (properties.cardSelectionAction.type === 'ExternalLink') {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardSelectionAction.parameters.target', {
          label: strings.LinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    if ((properties.cardSelectionAction.type as any) === 'TeamsExternalLink') {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardSelectionAction.parameters.teamsLink', {
          label: strings.TeamsLinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    actionsGroup.groupFields.push(PropertyPaneHorizontalRule());

    actionsGroup.groupFields.push(
      PropertyPaneToggle('cardButtonActions[0].isVisible', {
        label: strings.PrimaryButtonTogglePropertyLabel
      })
    );

    if (properties.cardButtonActions[0].isVisible) {
      actionsGroup.groupFields.push(
        PropertyPaneDropdown('cardButtonActions[0].action.type', {
          label: strings.PrimaryButtonActionPropertyLabel,
          options: [
            {
              key: 'QuickView',
              text: strings.ShowTheQuickViewLabel
            },
            {
              key: 'ExternalLink',
              text: strings.GoToALinkLabel
            },
            {
              key: 'TeamsExternalLink',
              text: strings.GoToATeamsAppLabel
            }
          ]
        })
      );
    }

    if (properties.cardButtonActions[0].isVisible) {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[0].title', {
          label: strings.PrimaryButtonTitlePropertyLabel
        })
      );
    }

    if (properties.cardButtonActions[0].isVisible && properties.cardButtonActions[0].action.type === 'ExternalLink') {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[0].action.parameters.target', {
          label: strings.PrimaryButtonLinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    // @todo: Review types for card button action type
    if (properties.cardButtonActions[0].isVisible && (properties.cardButtonActions[0].action.type as any) === 'TeamsExternalLink') {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[0].action.parameters.teamsLink', {
          label: strings.TeamsLinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    actionsGroup.groupFields.push(PropertyPaneHorizontalRule());

    actionsGroup.groupFields.push(
      PropertyPaneToggle('cardButtonActions[1].isVisible', {
        disabled: cardSize === 'Medium',
        label: strings.SecondaryButtonTogglePropertyLabel
      })
    );

    if (properties.cardButtonActions[1].isVisible && cardSize === 'Large') {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[1].title', {
          label: strings.SecondaryButtonTitlePropertyLabel
        })
      );
    }

    if (properties.cardButtonActions[1].isVisible && cardSize === 'Large') {
      actionsGroup.groupFields.push(
        PropertyPaneDropdown('cardButtonActions[1].action.type', {
          label: strings.SecondaryButtonActionPropertyLabel,
          options: [
            {
              key: 'QuickView',
              text: strings.ShowTheQuickViewLabel
            },
            {
              key: 'ExternalLink',
              text: strings.GoToALinkLabel
            },
            {
              key: 'TeamsExternalLink',
              text: strings.GoToATeamsAppLabel
            }
          ]
        })
      );
    }

    if (
      properties.cardButtonActions[1].isVisible &&
      cardSize === 'Large' &&
      properties.cardButtonActions[1].action.type === 'ExternalLink'
    ) {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[1].action.parameters.target', {
          label: strings.SecondaryButtonLinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    // @todo: Review types for card button action type
    if (
      properties.cardButtonActions[1].isVisible &&
      cardSize === 'Large' &&
      (properties.cardButtonActions[1].action.type as any) === 'TeamsExternalLink'
    ) {
      actionsGroup.groupFields.push(
        PropertyPaneTextField('cardButtonActions[1].action.parameters.teamsLink', {
          label: strings.TeamsLinkPropertyLabel,
          placeholder: 'https://'
        })
      );
    }

    //#endregion

    //#region Quick view data group
    const quickViewDataGroup: IPropertyPaneGroup = {
      groupName: strings.QuickViewDataGroupName,
      groupFields: []
    };

    if (properties._dynamicProperties.schema?.length > 0) {
      properties._dynamicProperties.schema.forEach((element) => {
        const fieldProperties = {
          label: element.label
        };

        let field: IPropertyPaneField<any>;

        if (element.type === 'PropertyPaneTextFieldMulti') {
          field = PropertyPaneTextField(`_dynamicProperties.values.${element.targetProperty}`, {
            ...fieldProperties,
            multiline: true
          });
        } else if (element.type === 'PropertyPaneToggle') {
          field = PropertyPaneToggle(`_dynamicProperties.values.${element.targetProperty}`, fieldProperties);
        } else {
          field = PropertyPaneTextField(`_dynamicProperties.values.${element.targetProperty}`, fieldProperties);
        }

        quickViewDataGroup.groupFields.push(field);
      });
    }
    //#endregion

    //#region Customization
    const customizationGroup: IPropertyPaneGroup = {
      groupName: strings.CustomizationGroupName,
      groupFields: []
    };

    customizationGroup.groupFields.push(
      PropertyFieldCodeEditor('quickViews[0].template', {
        initialValue: properties.quickViews[0].template,
        key: 'QuickViewsTemplateCodeEditor',
        label: strings.QuickViewTemplateJsonPropertyLabel,
        language: PropertyFieldCodeEditorLanguages.JSON,
        onPropertyChange: onPropertyPaneFieldChanged,
        panelTitle: strings.QuickViewTemplateJsonPropertyLabel,
        properties: properties
      })
    );

    customizationGroup.groupFields.push(
      PropertyFieldCollectionData('_dynamicProperties.schema', {
        key: 'DynamicPropertiesSchemaCollectionData',
        label: 'Data schema',
        panelHeader: 'Data schema',
        manageBtnLabel: 'Manage schema',
        value: properties._dynamicProperties.schema,
        fields: [
          {
            id: 'targetProperty',
            title: strings.DynamicPropertyTargetPropertyTitle,
            type: CustomCollectionFieldType.string,
            required: true
          },
          {
            id: 'label',
            title: strings.DynamicPropertyLabelTitle,
            type: CustomCollectionFieldType.string,
            required: true
          },
          {
            id: 'type',
            title: strings.DynamicPropertyTypeTitle,
            type: CustomCollectionFieldType.dropdown,
            required: true,
            options: [
              {
                key: 'PropertyPaneTextField',
                text: strings.SingleLineTextType
              },
              {
                key: 'PropertyPaneTextFieldMulti',
                text: strings.MultilineTextType
              },
              {
                key: 'PropertyPaneToggle',
                text: strings.ToggleType
              }
            ]
          },
          {
            id: 'source',
            title: strings.DynamicPropertySourceTitle,
            type: CustomCollectionFieldType.string,
            required: false
          }
        ]
      })
    );

    customizationGroup.groupFields.push(PropertyPaneHorizontalRule());

    customizationGroup.groupFields.push(
      PropertyPanePropertyEditor({
        webpart: adaptiveCardExtension,
        key: 'PropertyPanePropertyEditor'
      })
    );
    //#endregion

    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [sizeAndLayoutGroup, cardContentGroup, actionsGroup, quickViewDataGroup]
        },
        {
          displayGroupsAsAccordion: true,
          groups: [customizationGroup]
        }
      ]
    };
  }
}
