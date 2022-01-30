import {
  BaseAdaptiveCardExtension,
  ICardButton,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker.types';

import { CustomBaseBasicCardView } from './cardViews/CustomBaseBasicCardView';
import { CustomBaseImageCardView } from './cardViews/CustomBaseImageCardView';
import { CustomBasePrimaryTextCardView } from './cardViews/CustomBasePrimaryTextCardView';
import { CardIconSourceType, CardImageSourceType, CardTemplateType } from './CustomCardDesignerTypes';
import { QuickView } from './quickView/QuickView';

import type { CustomCardDesignerPropertyPane } from './CustomCardDesignerPropertyPane';

type CardButton = ICardButton & {
  isVisible: boolean;
};

type DynamicPropertySchema = {
  targetProperty: string;
  label: string;
  type: 'PropertyPaneTextField' | 'PropertyPaneTextFieldMulti' | 'PropertyPaneToggle';
  uniqueId: string;
  sortIdx: number;
};

export type CardDesignerAdaptiveCardExtensionProps = {
  templateType: string;
  title: string;
  description: string;
  cardIconSourceType: CardIconSourceType;
  primaryText: string;
  cardImageSourceType: CardImageSourceType;
  cardButtonActions: [CardButton] | [CardButton, CardButton];
  cardSelectionAction?: IQuickViewCardAction | IExternalLinkCardAction;
  iconProperty?: string;
  cardIconFilePickerResult?: IFilePickerResult;
  iconPicker?: string;
  cardImageFilePickerResult?: IFilePickerResult;
  imagePicker?: string;
  quickViews: [
    {
      data: string;
      template: string;
      id: string;
      displayName: string;
    }
  ];
  isQuickViewConfigured: boolean;
  currentQuickViewIndex: number;
  dataType: string;
  spRequestUrl: string;
  graphRequestUrl: string;
};

export type CustomCardDesignerAdaptiveCardExtensionProps = CardDesignerAdaptiveCardExtensionProps & {
  _dynamicProperties: {
    schema: DynamicPropertySchema[];
    values: { [k: string]: any };
  };
};

export type CustomCardDesignerAdaptiveCardExtensionState = {};

export const CUSTOM_BASE_BASIC_CARD_VIEW_REGISTRY_ID: string = 'CustomCardDesigner_BASE_BASIC_CARD_VIEW';
export const CUSTOM_BASE_IMAGE_CARD_VIEW_REGISTRY_ID: string = 'CustomCardDesigner_BASE_IMAGE_CARD_VIEW';
export const CUSTOM_BASE_PRIMARY_TEXT_CARD_VIEW_REGISTRY_ID: string = 'CustomCardDesigner_BASE_PRIMARY_TEXT_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'CustomCardDesigner_QUICK_VIEW';

export default class CustomCardDesignerAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  CustomCardDesignerAdaptiveCardExtensionProps,
  CustomCardDesignerAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CustomCardDesignerPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.cardNavigator.register(CUSTOM_BASE_BASIC_CARD_VIEW_REGISTRY_ID, () => new CustomBaseBasicCardView());
    this.cardNavigator.register(CUSTOM_BASE_IMAGE_CARD_VIEW_REGISTRY_ID, () => new CustomBaseImageCardView());
    this.cardNavigator.register(CUSTOM_BASE_PRIMARY_TEXT_CARD_VIEW_REGISTRY_ID, () => new CustomBasePrimaryTextCardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    if (this.properties.cardIconSourceType === CardIconSourceType.Icon) {
      return this.properties.iconPicker ?? '';
    }

    return this.properties.iconProperty ?? '';
  }

  protected async loadPropertyPaneResources(): Promise<void> {
    const PnPTelemetry = await import(
      /* webpackChunkName: 'CustomCardDesigner-pnp-teleemtry-js'*/
      '@pnp/telemetry-js'
    );

    const telemetry = PnPTelemetry.default.getInstance();
    telemetry.optOut();

    const component = await import(
      /* webpackChunkName: 'CustomCardDesigner-property-pane'*/
      './CustomCardDesignerPropertyPane'
    );

    this._deferredPropertyPane = new component.CustomCardDesignerPropertyPane();
  }

  protected renderCard(): string | undefined {
    if (this.properties.templateType === CardTemplateType.Heading) {
      return CUSTOM_BASE_BASIC_CARD_VIEW_REGISTRY_ID;
    } else if (this.properties.templateType === CardTemplateType.Image) {
      return CUSTOM_BASE_IMAGE_CARD_VIEW_REGISTRY_ID;
    } else if (this.properties.templateType === CardTemplateType.Description) {
      return CUSTOM_BASE_PRIMARY_TEXT_CARD_VIEW_REGISTRY_ID;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration({
      adaptiveCardExtension: this,
      cardSize: this.cardSize,
      context: this.context,
      onPropertyPaneFieldChanged: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties
    });
  }
}
