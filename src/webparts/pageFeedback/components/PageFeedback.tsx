import * as React from 'react';
import styles from './PageFeedback.module.scss';
import { IPageFeedbackProps } from './IPageFeedbackProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as AdaptiveCards from "adaptivecards";
import '../components/outlookstyles.css';

import { HttpClient } from "@microsoft/sp-http";
export default class PageFeedback extends React.Component<IPageFeedbackProps, {}> {
  public componentDidMount() {
  var adaptiveCard = new AdaptiveCards.AdaptiveCard();
 
    
const template={
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "type": "AdaptiveCard",
  "version": "1.3",
  "body": [
      {
          "type": "ColumnSet",
          "columns": [
              {
                  "type": "Column",
                  "width": 2,
                  "items": [
                      {
                          "type": "TextBlock",
                          "text": "Tell us what you think about this page",
                          "weight": "Bolder",
                          "size": "Medium",
                          "wrap": true
                      },
                      
                     
                      {
                          "type": "Container",
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Your feedback",
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "myFeedback",
                                  "isMultiline": true,
                                  "placeholder": "type your feedback here"
                              }
                          ]
                      },
                    
                  ]
              }
          ]
      }
  ],
  "actions": [
      {
          "type": "Action.Submit",
          "title": "Submit"
      }
  ]
};
adaptiveCard.onExecuteAction=(async (action:any) => {
  this.sendFeedbacktoTeams(action._processedData.myFeedback,this.props.context);
 
});
    adaptiveCard.parse(template);
    var renderedCard = adaptiveCard.render();
    document.getElementById("divFeedback").innerHTML = "";
    document.getElementById("divFeedback").appendChild(renderedCard);
  }
  public render(): React.ReactElement<IPageFeedbackProps> {
    return (
      <div className={ styles.pageFeedback }>
         <div id="divFeedback" />
      </div>
    );
  }
  public async sendFeedbacktoTeams(feedback:string,context:any) {
   
    const data = {
      "type": "message",
      "attachments": [
        {
          "contentType": "application/vnd.microsoft.card.adaptive",
          "content": {
          type: "AdaptiveCard",
          body: [
            {
              type: "ColumnSet",
              style: "emphasis",
              columns: [
                {
                  type: "Column",
                  items: [
                    {
                      type: "TextBlock",
                      size: "Large",
                      weight: "Bolder",
                      text: "Page feedback",
                      wrap: true,
                    },
                  ],
                  width: "stretch",
                  padding: "None",
                },
                {
                  type: "Column",
                  items: [
                    {
                      type: "TextBlock",
                      horizontalAlignment: "Right",
                      color: "Accent",
                      text:
                        `[View ${this.props.pageName}](${this.props.pageUrl})`,
                      wrap: true,
                    },
                  ],
                  width: "stretch",
                  padding: "None",
                },
              ],
              padding: "Default",
              spacing: "None",
            },
            
            {
              type: "Container",
              id: "7d00f965-40bb-9fc3-ff7b-a9b82a09ead4",
              padding: "Default",
              items: [
                {
                  type: "ColumnSet",
                  columns: [
                    {
                      type: "Column",
                      items: [
                        {
                          type: "Image",
                          style: "Person",
                          url:`https://m365404404.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${this.props.loginName}`,
                          size: "Small",
                        },
                      ],
                      width: "auto",
                      padding: "None",
                    },
                    {
                      type: "Column",
                      items: [
                        {
                          type: "TextBlock",
                          text: `${this.props.displayName}`,
                          wrap: true,
                        }
                      ],
                      width: "stretch",
                      padding: "None",
                    },
                  ],
                  spacing: "None",
                  padding: "None",
                },
              ],
              spacing: "None",
              separator: true,
            },
            {
              type: "Container",
              items: [
                {
                  type: "TextBlock",
                  text:
                    `*${feedback}*`,
                  wrap: true,
                },
              ],
              padding: "ExtraLarge",
              spacing: "None",
            }
         
          ],
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          version: "1.0",
          padding: "None",
        }
        }
    ]};
    const postURL = `https://outlook.office.com/webhook/2ff502de-743c-455a-be11-774a15121d84@c5489ec7-a322-45cf-a170-7ce0bdb538c5/IncomingWebhook/7bc3bb4fde2b429e9a6b8227581ff097/d21db31b-6fad-4550-9b18-c9f5a4f408d6`;
    return await context.httpClient.post(
      postURL,
      HttpClient.configurations.v1,
      {
        body: JSON.stringify(data),
        mode: "no-cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );
  }

}
