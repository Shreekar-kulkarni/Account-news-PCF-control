import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as AdaptiveCards from "adaptivecards";

interface newsDetails {
    newsTitle: string,
    newsDescription: string,
    imgUrl: string,
    sourceUrl: string,
    publishTime: string,
    newsSource: string,
    sentimentScore: number
}

export class AccountNewsControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _cardContainer: HTMLDivElement;
    private _positiveSentimentURL: string = "https://org64ec1af3.crm.dynamics.com/WebResources/sh_Positive_sentiment_image";
    private _veryPositiveSentimentURL: string = "https://org64ec1af3.crm.dynamics.com/WebResources/sh_Very_positive_sentiment_image";
    private _neutralSentimentURL: string = "https://org64ec1af3.crm.dynamics.com/WebResources/sh_Neutral_sentiment_image";
    private _negativeSentimentURL: string = "https://org64ec1af3.crm.dynamics.com/WebResources/sh_Negative_sentiment_image";
    private _veryNegativeSentimentURL: string = "https://org64ec1af3.crm.dynamics.com/WebResources/sh_Very_negative_sentiment_image";
    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        console.log("Account news Component Called");
        this._container = container;
        this._context = context;

        // this._textInputElement = document.createElement("input");
        // this._textInputElement.type = "text";
        // this._textInputElement.addEventListener("change", this.getKeywordsFromText.bind(this));
        // this._container.appendChild(this._textInputElement);
        let accountName = this._context.parameters.AccountName.raw || " ";
        if(accountName !=" "){
            this.getNews(accountName);
        }else{
            let h1 = document.createElement("h1");
            h1.innerHTML = "Please enter the name for account";
            this._container.appendChild(h1);
        }
    }

    public getNews(query: string) {
        let apiKey = this._context.parameters.ApiKey.raw || " ";

        fetch(`https://api.bing.microsoft.com/v7.0/news/search?q=${query}&count=100&sortBy=Date&originalImg=true`,{
            method: 'GET',
            headers:{
                'Ocp-Apim-Subscription-Key': apiKey,
            }
        })
            .then((response) => {
                return response.json();
            })
            .then((newsJson) => {
                console.log(newsJson);
                this.createCard(newsJson);
            });
    }

    public createCard(newsJson: any) {
        //console.log(newsJson);
        newsJson.value.forEach((element: any) => {
            console.log(element);
            let newsDetail: newsDetails = {
                newsTitle: element.name,
                newsDescription: element.description,
                sourceUrl: element.url,
                imgUrl:(element.image !=null)? element.image.contentUrl : "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS1MiBflN17NfMjCKamD-u31XZFSWnelPtYKQ&usqp=CAU",
                publishTime: element.datePublished,
                newsSource: element.provider[0].name,
                sentimentScore: element.sentiment
            };
            let card = this.getCard(newsDetail);
            let adaptiveCard = new AdaptiveCards.AdaptiveCard();
            adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
                fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
            });
            adaptiveCard.onExecuteAction = (action) => {
                window.open(newsDetail.sourceUrl, '_blank');
            }
            adaptiveCard.parse(card);

            let renderedCard: any = adaptiveCard.render();
            let button = document.createElement("button");
            button.innerHTML = "View full story";
            button.onclick = () => {
                window.open(newsDetail.sourceUrl, '_blank');
            }
            //button.setAttribute("class", "button");
            this._cardContainer = document.createElement("div");
            this._cardContainer.appendChild(renderedCard);
            this._cardContainer.appendChild(button);
            this._container.appendChild(this._cardContainer);
        });
    }

    public getCard(newsJson: newsDetails) {
        let publishedDate = new Date(newsJson.publishTime).toLocaleDateString("en-GB", { day: 'numeric', month: 'short', year: 'numeric' });
        //console.log(publishedDate);
        let newsSource = newsJson.newsSource;
        //console.log(newsSource);
        let sentimentScore = newsJson.sentimentScore;
        let sentimentURL = "";
        if (sentimentScore <= -0.4) {
            sentimentURL = this._veryNegativeSentimentURL;
        }
        else if (sentimentScore > -0.4 && sentimentScore <= 0.1) {
            sentimentURL = this._negativeSentimentURL;
        }
        else if (sentimentScore > -0.1 && sentimentScore < 0.1) {
            sentimentURL = this._neutralSentimentURL;
        }
        else if (sentimentScore >= 0.1 && sentimentScore < 0.5) {
            sentimentURL = this._positiveSentimentURL;
        }
        else {
            sentimentURL = this._veryPositiveSentimentURL;
        }
        let card = {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.6",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "auto",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "url": newsJson.imgUrl,
                                            "size": "medium",
                                            "height": "150px",
                                            "width": "150px",
                                            "selectAction": {
                                                "type": "Action.OpenUrl",
                                                "title": "View full story",
                                            },
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "stretch",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "text": newsJson.newsTitle,
                                            "weight": "bolder",
                                            "size": "large",
                                            "wrap": true
                                        },
                                        {
                                            "type": "RichTextBlock",
                                            "inlines": [
                                                {
                                                    "type": "TextRun",
                                                    "text": publishedDate,
                                                    "weight": "bolder"
                                                },
                                                {
                                                    "type": "TextRun",
                                                    "text": " Source: "
                                                },
                                                {
                                                    "type": "TextRun",
                                                    "text": newsSource,
                                                    "weight": "bolder"
                                                }
                                            ]
                                        },
                                        {
                                            "type": "TextBlock",
                                            "text": newsJson.newsDescription,
                                            "wrap": true,
                                            "maxLines": 2
                                        },
                                        {
                                            "type": "Image",
                                            "url": sentimentURL,
                                            "height": "30px"
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
        return card;
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}
