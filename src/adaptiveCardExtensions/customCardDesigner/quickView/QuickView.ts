import { BaseAdaptiveCardView, ISPFxAdaptiveCard } from '@microsoft/sp-adaptive-card-extension-base';
import {
  CustomCardDesignerAdaptiveCardExtensionProps,
  CustomCardDesignerAdaptiveCardExtensionState
} from '../CustomCardDesignerAdaptiveCardExtension';

export type QuickViewData = {
  [k: string]: any;
  $properties: CustomCardDesignerAdaptiveCardExtensionProps;
  $state: CustomCardDesignerAdaptiveCardExtensionState;
};

export class QuickView extends BaseAdaptiveCardView<
  CustomCardDesignerAdaptiveCardExtensionProps,
  CustomCardDesignerAdaptiveCardExtensionState,
  QuickViewData
> {
  public get data(): QuickViewData {
    return {
      ...this.properties._dynamicProperties.values,
      $properties: this.properties,
      $state: this.state
    };
  }

  public get template(): ISPFxAdaptiveCard {
    const template = this.properties.quickViews[0].template;
    const json = JSON.parse(template);

    return json;
  }
}
