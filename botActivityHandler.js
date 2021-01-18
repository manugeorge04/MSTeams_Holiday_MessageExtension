// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
    }


    getAdaptiveCard (holidaylist, location) {        
        let adaptiveCardAttachment = {
            contentType: "application/vnd.microsoft.card.adaptive",            
            content: {
              type: "AdaptiveCard",
              version: "1.0",
              body: [
                {
                    type: "RichTextBlock",
                    inlines: [
                        {
                            type: "TextRun",
                            text: `Holidays - ${location}`,
                            size: "Large",
                            weight: "Bolder"
                        }
                    ]
                },
                {
                    type: "Container",
                    items: [
                        {
                            type: "ColumnSet",
                            columns: [
                                {
                                    type: "Column",
                                    width: "stretch",
                                    items: [
                                        {
                                            type: "TextBlock",
                                            text: "Occasion",
                                            size: "Medium",   
                                            weight: "Bolder",                                 
                                            wrap: true
                                        }
                                    ]
                                },
                                {
                                    type: "Column",
                                    width: "stretch",
                                    items: [
                                        {
                                            type: "TextBlock",
                                            text: "Date",
                                            weight: "Bolder",
                                            size: "Medium",
                                            wrap: true
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }//,
                //result of map function to be added here
              ]                          
            }
          }

        const adaptiveCardBody = holidaylist.map((holiday) => {
            return{
                type: "ColumnSet",
                columns: [
                    {
                        type: "Column",
                        width: "stretch",
                        items: [
                            {
                                type: "TextBlock",
                                text: holiday.items[0],
                                wrap: true
                            }
                        ]
                    },
                    {
                        type: "Column",
                        width: "stretch",
                        items: [
                            {
                                type: "TextBlock",
                                text: holiday.items[1],
                                wrap: true
                            }
                        ]
                    }
                ]
            }
        })

        const updatedBody = adaptiveCardAttachment.content.body.concat(adaptiveCardBody)

        adaptiveCardAttachment.content.body = updatedBody

        return adaptiveCardAttachment
    }

    handleTeamsMessagingExtensionFetchTask(context, action) {        
        return {
        task: {
            type: 'continue',
            value: {
            width: 450,
            height: 800,
            title: 'Holiday List',
            url: 'https://holidaylist.azurewebsites.net/',
            fallbackUrl: 'https://holidaylist.azurewebsites.net/'
            }
        }
        };        
    }


    handleTeamsMessagingExtensionSubmitAction(context, action) {                
        const holidaylist = action.data.holidayList
        const location = action.data.location        

        const adaptiveCardBody = this.getAdaptiveCard(holidaylist, location)
        const attachment = adaptiveCardBody      
        
        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [
                    attachment
                ]
            }
        };   
    }   
}

module.exports.BotActivityHandler = BotActivityHandler;