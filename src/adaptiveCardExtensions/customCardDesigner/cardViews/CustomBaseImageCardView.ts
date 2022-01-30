import {
  BaseImageCardView,
  ICardButton,
  IExternalLinkCardAction,
  IImageCardParameters,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  CustomCardDesignerAdaptiveCardExtensionProps,
  CustomCardDesignerAdaptiveCardExtensionState
} from '../CustomCardDesignerAdaptiveCardExtension';

export class CustomBaseImageCardView extends BaseImageCardView<
  CustomCardDesignerAdaptiveCardExtensionProps,
  CustomCardDesignerAdaptiveCardExtensionState
> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return this.properties.cardButtonActions;
  }

  public get data(): IImageCardParameters {
    return {
      imageUrl: this.properties.imagePicker,
      primaryText: this.properties.primaryText
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    const cardSelectionAction: IQuickViewCardAction | IExternalLinkCardAction | undefined = this.properties.cardSelectionAction;

    if ((cardSelectionAction.type as any) === 'TeamsExternalLink') {
      cardSelectionAction.type = 'ExternalLink';
      (cardSelectionAction as IExternalLinkCardAction).parameters.isTeamsDeepLink = true;
    }

    return this.properties.cardSelectionAction;
  }
}
