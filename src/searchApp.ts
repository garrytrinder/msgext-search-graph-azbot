import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  CloudAdapter,
  AttachmentLayoutTypes,
  CardImage,
} from "botbuilder";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import config from "./config";

const listFields = [
  "fields/Title",
  "fields/RetailCategory",
  "fields/Specguide",
  "fields/PhotoSubmission",
  "fields/CustomerRating",
  "fields/ReleaseDate"
];

export class SearchApp extends TeamsActivityHandler {
constructor() {
    super();
  }

  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {
    const magicCode = query.state && Number.isInteger(Number(query.state)) ? query.state : '';
    const userTokenClient = context.turnState.get((context.adapter as CloudAdapter).UserTokenClientKey);
    const tokenResponse = await userTokenClient.getUserToken(context.activity.from.id, config.oauthConnectionName, context.activity.channelId, magicCode);

    if (!tokenResponse || !tokenResponse.token) {
      const { signInLink } = await userTokenClient.getSignInResource(config.oauthConnectionName, context.activity);

      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                value: signInLink,
                title: 'Bot Service OAuth'
              }
            ]
          }
        }
      };
    }

    const graphClient = Client.init({ authProvider: (done) => { done(null, tokenResponse.token); } });

    const { sharepointIds } = await graphClient.api(`/sites/${config.spoHostname}:/${config.spoSiteUrl}`).select("sharepointIds").get();
    const { value: items } = await graphClient.api(`/sites/${sharepointIds.siteId}/lists/Products/items?expand=fields&select=${listFields.join(",")}&$filter=startswith(fields/Title,'${query.parameters[0].value}')`).get();
    const { value: drives } = await graphClient.api(`sites/${sharepointIds.siteId}/drives`).select(["id", "name"]).get();
    const drive = drives.find(drive => drive.name === "Product Imagery");

    const attachments = [];
    await Promise.all(items.map(async (item) => {
      const { PhotoSubmission: photoUrl, Title, RetailCategory } = item.fields;
      const fileName = photoUrl.split("/").reverse()[0];
      const driveItem = await graphClient.api(`sites/${sharepointIds.siteId}/drives/${drive.id}/root:/${fileName}`).get();
      const content = await graphClient.api(`sites/${sharepointIds.siteId}/drives/${drive.id}/items/${driveItem.id}/content`).responseType(ResponseType.ARRAYBUFFER).get();
      const cardImages: CardImage[] = [{ url: `data:${driveItem.file.mimeType};base64,${Buffer.from(content).toString('base64')}`, alt: Title }]
      const card = CardFactory.thumbnailCard(Title, RetailCategory, cardImages);
      attachments.push(card);
    }));

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: AttachmentLayoutTypes.List,
        attachments,
      },
    } as MessagingExtensionResponse;
  }
}
