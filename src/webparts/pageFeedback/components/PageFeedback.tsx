import * as React from "react";
import styles from "./PageFeedback.module.scss";
import { IPageFeedbackProps } from "./IPageFeedbackProps";
import * as AdaptiveCards from "adaptivecards";
import "../components/outlookstyles.css";
import { MSGraphClient, HttpClient } from "@microsoft/sp-http";

export interface IPageFeedbackState {
  jobTitle;
  mobilePhone;
  officeLocation;
  sent:boolean;
}
export default class PageFeedback extends React.Component<IPageFeedbackProps,IPageFeedbackState> {
  constructor(props: IPageFeedbackProps, state: IPageFeedbackState) {
    super(props);
    this.state = {
      jobTitle: "",
      mobilePhone: "",
      officeLocation: "",
      sent:false
    };
  }

  //graph call to get more info of the user
  protected getUserProfile(): void {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        return client
          .api("me")
          .version("v1.0")
          .select("jobTitle,mobilePhone,officeLocation")
          .get()
          .then((res) => {
            this.setState({
              jobTitle: res.jobTitle,
              mobilePhone: res.mobilePhone,
              officeLocation: res.officeLocation,
            });
          });
      });
  }
  public componentDidMount() {
    this.getUserProfile();
    var acFeedback = new AdaptiveCards.AdaptiveCard();
    const templateFeedback = {
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.3",
      body: [
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              width: 2,
              items: [
                {
                  type: "TextBlock",
                  text: "Tell us what you think about this page",
                  weight: "Bolder",
                  size: "Medium",
                  wrap: true,
                },

                {
                  type: "Container",
                  items: [
                    {
                      type: "TextBlock",
                      text: "Your feedback",
                      wrap: true,
                    },
                    {
                      type: "Input.Text",
                      id: "myFeedback",
                      isMultiline: true,
                      placeholder: "type your feedback here",
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Submit",
        },
      ],
    };
    //adaptive card from feedback form submit action handler
    acFeedback.onExecuteAction = async (action: any) => {
      this.sendFeedbacktoTeams(
        action._processedData.myFeedback,
        this.props.context
      );
      this.setState({sent:true});
    };
    acFeedback.parse(templateFeedback);
    var renderedCard = acFeedback.render(); //render card for feedbac
    document.getElementById("divFeedback").innerHTML = "";
    document.getElementById("divFeedback").appendChild(renderedCard);
  }

  //render feedback react component
  public render(): React.ReactElement<IPageFeedbackProps> {
    return (
      <div className={styles.pageFeedback}>
       {this.state.sent?
       <div id="divFeedbackSent" >Thank you for your feedback 👍🏽</div>:
       <div id="divFeedback" />
      }
      </div>
    );
  }

//function to send feedback received notification to the channel using incoming webhooks
  public async sendFeedbacktoTeams(feedback: string, context: any) {
    const card = {
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: {
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
                        text: `[View ${this.props.pageName}](${this.props.pageUrl})`,
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
                padding: "Large",
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
                            url: `https://m365404404.sharepoint.com/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${this.props.loginName}`,
                            size: "Medium",
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
                          },
                          {
                            type: "TextBlock",
                            spacing: "None",
                            isSubtle: true,
                            text: `${this.state.jobTitle}`,
                            wrap: true,
                            size: "Small",
                          },
                          {
                            type: "TextBlock",
                            spacing: "None",
                            isSubtle: true,
                            text: `${this.state.mobilePhone}`,
                            wrap: true,
                            size: "Small",
                          },
                          {
                            type: "TextBlock",
                            spacing: "None",
                            isSubtle: true,
                            text: `${this.state.officeLocation}`,
                            wrap: true,
                            size: "Small",
                          },
                        ],
                        width: "stretch",
                        padding: "None",
                      },
                    ],
                    spacing: "Large",
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
                    text: `"*${feedback}*"`,
                    wrap: true,
                  },
                ],
                padding: "ExtraLarge",
                spacing: "None",
              },
            ],
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            version: "1.0",
            padding: "None",
          },
        },
      ],
    };
    return await context.httpClient.post(
      this.props.connectorUrl,
      HttpClient.configurations.v1,
      {
        body: JSON.stringify(card),
        mode: "no-cors",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      }
    );
  }
}
