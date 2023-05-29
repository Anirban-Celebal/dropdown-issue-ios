const { TeamsActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const card1 = {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
        {
            type: 'ActionSet',
            actions: [
                {
                    type: 'Action.Submit',
                    title: 'Open Task Module',
                    data: {
                        action: 'test',
                        msteams: {
                            type: 'task/fetch'
                        }
                    }
                }
            ]
        }
    ]
};

const card2 = {
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
        {
            "type": "TextBlock",
            "text": "Main heading",
            "size": "ExtraLarge",
            "weight": "Bolder",
            "wrap": true
        },
        {
            "type": "TextBlock",
            "text": "Sub heading",
            "wrap": true,
            "weight": "Default",
            "spacing": "Small",
            "isSubtle": true,
            "color": "Default",
            "size": "Small"
        },
        {
            "type": "Container",
            "spacing": "Large",
            "separator": true,
            "items": [
                {
                    "type": "Container",
                    "spacing": "Medium",
                    "separator": true,
                    "items": [
                        {
                            "type": "Container",
                            "items": [
                                {
                                    "type": "Container",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "ColumnSet",
                                            "columns": [
                                                {
                                                    "type": "Column",
                                                    "width": 95,
                                                    "items": [
                                                        {
                                                            "type": "RichTextBlock",
                                                            "inlines": [
                                                                {
                                                                    "type": "TextRun",
                                                                    "text": "Melanie Griff",
                                                                    "weight": "Bolder"
                                                                },
                                                                {
                                                                    "type": "TextRun",
                                                                    "text": " (Executive Director)",
                                                                    "color": "Accent"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "TextBlock",
                                                            "text": "Chief Executive",
                                                            "wrap": true,
                                                            "spacing": "Small"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Column",
                                                    "id": "expand0",
                                                    "isVisible": true,
                                                    "width": 5,
                                                    "verticalContentAlignment": "Center",
                                                    "items": [
                                                        {
                                                            "type": "Image",
                                                            "url": "https://azamawa4.azurewebsites.net/static/images/arrow_down.png",
                                                            "width": "16px",
                                                            "height": "16px"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Column",
                                                    "id": "collapse0",
                                                    "width": 5,
                                                    "isVisible": false,
                                                    "verticalContentAlignment": "Center",
                                                    "items": [
                                                        {
                                                            "type": "Image",
                                                            "url": "https://azamawa4.azurewebsites.net/static/images/arrow_up.png",
                                                            "width": "16px",
                                                            "height": "16px"
                                                        }
                                                    ]
                                                }
                                            ]
                                        },
                                        {
                                            "type": "Container",
                                            "id": "data0",
                                            "isVisible": false,
                                            "spacing": "Large",
                                            "items": [
                                                {
                                                    "type": "ColumnSet",
                                                    "isVisible": true,
                                                    "separator": true,
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "verticalContentAlignment": "Center",
                                                            "items": [
                                                                {
                                                                    "type": "Image",
                                                                    "url": "https://azamawa4.azurewebsites.net/static/images/icon_call.png",
                                                                    "width": "25px",
                                                                    "height": "25px"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "stretch",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "[+7707183770](+7707183770)",
                                                                    "wrap": true
                                                                }
                                                            ],
                                                            "verticalContentAlignment": "Center",
                                                            "selectAction": {
                                                                "type": "Action.OpenUrl",
                                                                "url": "tel:+7707183770"
                                                            }
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "ColumnSet",
                                                    "isVisible": false,
                                                    "separator": true,
                                                    "spacing": "Large",
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "verticalContentAlignment": "Center",
                                                            "items": [
                                                                {
                                                                    "type": "Image",
                                                                    "url": "https://azamawa4.azurewebsites.net/static/images/icon_email.png",
                                                                    "width": "25px",
                                                                    "height": "25px"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "stretch",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "[noemail@gmail.com](mailto:noemail@gmail.com)",
                                                                    "wrap": true
                                                                }
                                                            ],
                                                            "verticalContentAlignment": "Center",
                                                            "selectAction": {
                                                                "type": "Action.OpenUrl",
                                                                "url": "mailto:noemail@gmail.com"
                                                            }
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ],
                            "spacing": "Medium",
                            "separator": true,
                            "selectAction": {
                                "type": "Action.ToggleVisibility",
                                "targetElements": [
                                    "expand0",
                                    "collapse0",
                                    "data0"
                                ]
                            }
                        }
                    ]
                },
                {
                    "type": "Container",
                    "spacing": "Medium",
                    "separator": true,
                    "items": [
                        {
                            "type": "Container",
                            "items": [
                                {
                                    "type": "Container",
                                    "style": "emphasis",
                                    "items": [
                                        {
                                            "type": "ColumnSet",
                                            "columns": [
                                                {
                                                    "type": "Column",
                                                    "width": 95,
                                                    "items": [
                                                        {
                                                            "type": "RichTextBlock",
                                                            "inlines": [
                                                                {
                                                                    "type": "TextRun",
                                                                    "text": "Leigha Falls",
                                                                    "weight": "Bolder"
                                                                },
                                                                {
                                                                    "type": "TextRun",
                                                                    "text": " (Chief Technical Officer)",
                                                                    "color": "Accent"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "TextBlock",
                                                            "text": "Executive",
                                                            "wrap": true,
                                                            "spacing": "Small"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Column",
                                                    "id": "expand1",
                                                    "isVisible": true,
                                                    "width": 5,
                                                    "verticalContentAlignment": "Center",
                                                    "items": [
                                                        {
                                                            "type": "Image",
                                                            "url": "https://azamawa4.azurewebsites.net/static/images/arrow_down.png",
                                                            "width": "16px",
                                                            "height": "16px"
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "Column",
                                                    "id": "collapse1",
                                                    "width": 5,
                                                    "isVisible": false,
                                                    "verticalContentAlignment": "Center",
                                                    "items": [
                                                        {
                                                            "type": "Image",
                                                            "url": "https://azamawa4.azurewebsites.net/static/images/arrow_up.png",
                                                            "width": "16px",
                                                            "height": "16px"
                                                        }
                                                    ]
                                                }
                                            ]
                                        },
                                        {
                                            "type": "Container",
                                            "id": "data1",
                                            "isVisible": false,
                                            "spacing": "Large",
                                            "items": [
                                                {
                                                    "type": "ColumnSet",
                                                    "isVisible": true,
                                                    "separator": true,
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "verticalContentAlignment": "Center",
                                                            "items": [
                                                                {
                                                                    "type": "Image",
                                                                    "url": "https://azamawa4.azurewebsites.net/static/images/icon_call.png",
                                                                    "width": "25px",
                                                                    "height": "25px"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "stretch",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "[+7709111157](+7709111157)",
                                                                    "wrap": true
                                                                }
                                                            ],
                                                            "verticalContentAlignment": "Center",
                                                            "selectAction": {
                                                                "type": "Action.OpenUrl",
                                                                "url": "tel:+7709111157"
                                                            }
                                                        }
                                                    ]
                                                },
                                                {
                                                    "type": "ColumnSet",
                                                    "isVisible": false,
                                                    "separator": true,
                                                    "spacing": "Large",
                                                    "columns": [
                                                        {
                                                            "type": "Column",
                                                            "width": "auto",
                                                            "verticalContentAlignment": "Center",
                                                            "items": [
                                                                {
                                                                    "type": "Image",
                                                                    "url": "https://azamawa4.azurewebsites.net/static/images/icon_email.png",
                                                                    "width": "25px",
                                                                    "height": "25px"
                                                                }
                                                            ]
                                                        },
                                                        {
                                                            "type": "Column",
                                                            "width": "stretch",
                                                            "items": [
                                                                {
                                                                    "type": "TextBlock",
                                                                    "text": "[noemail@gmail.com](mailto:noemail@gmail.com)",
                                                                    "wrap": true
                                                                }
                                                            ],
                                                            "verticalContentAlignment": "Center",
                                                            "selectAction": {
                                                                "type": "Action.OpenUrl",
                                                                "url": "mailto:noemail@gmail.com"
                                                            }
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ],
                            "spacing": "Medium",
                            "separator": true,
                            "selectAction": {
                                "type": "Action.ToggleVisibility",
                                "targetElements": [
                                    "expand1",
                                    "collapse1",
                                    "data1"
                                ]
                            }
                        }
                    ]
                }
            ]
        }
    ]
}

class EchoBot extends TeamsActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card1)] });
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async handleTeamsTaskModuleFetch() {
        try {
            return {
                task: {
                    type: 'continue',
                    value: {
                        height: 'medium',
                        width: 'medium',
                        title: 'Test Module',
                        card: CardFactory.adaptiveCard(card2)
                    }
                }
            };
        } catch (error) {
            return null;
        }
    }
}

module.exports.EchoBot = EchoBot;
