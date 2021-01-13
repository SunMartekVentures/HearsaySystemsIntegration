'use strict';

import axios from 'axios';
import express = require("express");
import Utils from './Utils';
import xml2js = require("xml2js");

export default class SfmcApiHelper
{
    // Instance variables 
    private _deExternalKey = "Org_Setup";
    private _sfmcDataExtensionApiUrl = "https://mc4f63jqqhfc51yw6d1h0n1ns1.rest.marketingcloudapis.com/hub/v1/dataevents/key:" + this._deExternalKey + "/rowset";
    private _oauthToken = "";
	private FolderID='';
	private ParentFolderID = '';
	private StatusCode = '';
	private DataExtensionName = '';
	private Hearsay_Org_ID = '';
	private validateStatus = '';
	private validateDEName = '';
	private isValidated = '';
	//private xmlDoc = '';
    
    
    /**
     * getOAuthAccessToken: POSTs to SFMC Auth URL to get an OAuth access token with the given ClientId and ClientSecret
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.noversion.mc-getting-started.meta/mc-getting-started/get-access-token.htm
     * 
     */
    public getOAuthAccessToken(clientId: string, clientSecret: string) : Promise<any>
    {
        let self = this;
        Utils.logInfo("getOAuthAccessToken called.");
        Utils.logInfo("Using specified ClientID and ClientSecret to get OAuth token...");

        let headers = {
            'Content-Type': 'application/json',
        };

        let postBody = {
            'grant_type': 'client_credentials',
            'client_id': clientId,
            'client_secret': clientSecret
        };

        return self.getOAuthTokenHelper(headers, postBody);
    }

    /**
     * getOAuthAccessTokenFromRefreshToken: POSTs to SFMC Auth URL to get an OAuth access token with the given refreshToken
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.noversion.mc-getting-started.meta/mc-getting-started/get-access-token.htm
     * 
     */
    public getOAuthAccessTokenFromRefreshToken(clientId: string, clientSecret: string, refreshToken: string) : Promise<any>
    {
        let self = this;
        Utils.logInfo("getOAuthAccessTokenFromRefreshToken called.");
        Utils.logInfo("Getting OAuth Access Token with refreshToken: " + refreshToken);
        
        let headers = {
            'Content-Type': 'application/json',
        };

        let postBody = {
            'clientId': clientId,
            'clientSecret': clientSecret,
            'refreshToken': refreshToken
        };

        return self.getOAuthTokenHelper(headers, postBody);
    }

    /**
     * getOAuthTokenHelper: Helper method to POST the given header & body to the SFMC Auth endpoint
     * 
     */
    public getOAuthTokenHelper(headers : any, postBody: any) : Promise<any>
    {
        return new Promise<any>((resolve, reject) =>
        {
            Utils.logInfo("Entered to the method...");
            // POST to Marketing Cloud REST Auth service and get back an OAuth access token.
            let sfmcAuthServiceApiUrl = "https://mc4f63jqqhfc51yw6d1h0n1ns1.auth.marketingcloudapis.com/v2/token";
            Utils.logInfo("oauth token is called, waiting for status...");
            axios.post(sfmcAuthServiceApiUrl, postBody, {"headers" : headers})            
            .then((response : any) => {
                // success
                Utils.logInfo("Success, got auth token from MC...");
                let accessToken = response.data.access_token;
                Utils.logInfo("oauth token..." + accessToken);
                let bearer = response.data.token_type;
                Utils.logInfo("Bearer..." + bearer);
                let tokenExpiry = response.data.expires_in;
                Utils.logInfo("tokenExpiry..." + tokenExpiry);
                
				this._oauthToken = response.data.access_token;
				Utils.logInfo("Storing the accesstoken in a object's variable "+ this._oauthToken);
                //tokenExpiry.setSeconds(tokenExpiry.getSeconds() + response.data.expires_in);
                Utils.logInfo("Got OAuth token: " + accessToken + ", expires = " +  tokenExpiry);

                resolve(
                {
                    oauthAccessToken: accessToken,
                    oauthAccessTokenExpiry: tokenExpiry,
                    status: response.status,
                    statusText: response.statusText + "\n" + Utils.prettyPrintJson(JSON.stringify(response.data))
                });
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error getting OAuth Access Token.";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
        });
    }
	
	public getCategoryID(req: express.Request, res: express.Response)
    {
		Utils.logInfo("Get Category Method: " + this._oauthToken);
		let self = this;	
			
		Utils.logInfo("request body = " + JSON.stringify(req.query.TemplateName));
		let TemplateName = JSON.stringify(req.query.TemplateName);
		if (this._oauthToken!= "")
        {
            //Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.getCategoryIDHelper(this._oauthToken, TemplateName)
            .then((result) => {
                res.status(result.status).send(result.statusText);
            })
            .catch((err) => {
                res.status(500).send(err);
            });
        }
        else
        {
            // error
            let errorMsg = "OAuth Access Token *not* found in session.\nPlease complete previous demo step\nto get an OAuth Access Token."; 
            Utils.logError(errorMsg);
            res.status(500).send(errorMsg);
        }
	}
	public getCategoryIDHelper(oauthAccessToken: string, TemplateName : string) : Promise<any>
	{
		
		//Utils.logInfo("Template Name " );
		
		let validateName = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Retrieve</a:Action>'
+'        <a:To s:mustUnderstand="1">https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'   <RetrieveRequestMsg xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'      <RetrieveRequest>'
+'         <ObjectType>DataExtension</ObjectType>'
+'         <Properties>ObjectID</Properties>'
+'         <Properties>CustomerKey</Properties>'
+'         <Properties>Name</Properties>'
+'         <Properties>IsSendable</Properties>'
+'         <Properties>SendableSubscriberField.Name</Properties>'
+'         <Filter xsi:type="SimpleFilterPart">'
+'            <Property>Name</Property>'
+'            <SimpleOperator>equals</SimpleOperator>'
+'            <Value>'+JSON.parse(TemplateName)+'</Value>'
+'         </Filter>'
+'      </RetrieveRequest>'
+'   </RetrieveRequestMsg>'
+'    </s:Body>'
+'</s:Envelope>';


		let soapMessage = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Retrieve</a:Action>'
+'        <a:To s:mustUnderstand="1">https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'        <RetrieveRequestMsg xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'            <RetrieveRequest>'
+'                <ObjectType>DataFolder</ObjectType>'
+'                <Properties>ID</Properties>'
+'                <Properties>CustomerKey</Properties>'
+'                <Properties>Name</Properties>'
+'                <Properties>ParentFolder.ID</Properties>'
+'                <Properties>ParentFolder.Name</Properties>'
+'                <Filter xsi:type="SimpleFilterPart">'
+'                    <Property>Name</Property>'
+'                    <SimpleOperator>equals</SimpleOperator>'
+'                    <Value>Hearsay Integrations</Value>'
+'                </Filter>'
+'            </RetrieveRequest>'
+'        </RetrieveRequestMsg>'
+'    </s:Body>'
+'</s:Envelope>';
				
	return new Promise<any>((resolve, reject) =>
		{
			let headers = {
                'Content-Type': 'text/xml',
                'SOAPAction': 'Retrieve'
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios({
				method: 'post',
				url: 'https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx',
				data: soapMessage,
				headers: {'Content-Type': 'text/xml'}							
				})            
				.then((response: any) => {
                Utils.logInfo(response.data);
                var extractedData = "";
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				this.FolderID = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ID'][0];
				Utils.logInfo('Folder ID : ' + this.FolderID);
				this.ParentFolderID = JSON.stringify(result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ParentFolder'][0]);
				Utils.logInfo('Parent Folder ID : ' + this.ParentFolderID);
				this.ValidationForDataExtName(validateName);
				});

			})
			.catch((error: any) => {
						// error
						let errorMsg = "Error loading sample data. POST response from Marketing Cloud:";
						errorMsg += "\nMessage: " + error.message;
						errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
						errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
						Utils.logError(errorMsg);

										reject(errorMsg);
									});
			
        });
	}
	
	private ValidationForDataExtName(ValidationBody : any) : Promise<any>
	{
		
			Utils.logInfo("Validation Body : "+ ValidationBody);
			
			//let headers = {
                //'Content-Type': 'application/json',
                //'Authorization': 'Bearer ' + this._oauthToken
            //};
		
		return new Promise<any>((resolve, reject) =>
        {
			Utils.logInfo("Ahpppaaaddaa, Method call aaiduchu");
			 axios({
				method: 'post',
				url: 'https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx',
				data: ValidationBody,
				headers: {'Content-Type': 'text/xml'}							
				}) 
            .then((response: any) => {
                // success
                Utils.logInfo("Validation Successful \n\n" + response.data);
				
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				this.validateStatus = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['OverallStatus'][0];
				Utils.logInfo('Validation Status : ' + this.validateStatus);
				this.validateDEName = JSON.stringify(result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['Name'][0]);
				Utils.logInfo('Validated Data Extension Name : ' + this.validateDEName);
					if(this.validateStatus =='OK' && this.validateDEName!=""){
						this.isValidated = 'true';
					}
					else{
						this.isValidated = 'false';
					}
                
            });
			})
            .catch((error: any) => {
                // error
                let errorMsg = "Error loading sample data. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
			
	})
	}
	

    /**
     * loadData: called by the GET handlers for /apidemoloaddata and /appdemoloaddata
     * 
     */
    public loadData(req: express.Request, res: express.Response)
    {
        Utils.logInfo("request body = " + JSON.stringify(req.body));
        let self = this;
        let sessionId = req.session.id;
        Utils.logInfo("loadData entered. SessionId = " + sessionId);
	    //Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
		Utils.logInfo("Getting the accesstoken in a object's variable "+ this._oauthToken);

        if (this._oauthToken!= "")
        {
            //Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.loadDataHelper(this._oauthToken, JSON.stringify(req.body))
            .then((result) => {
                res.status(result.status).send(result.statusText);
            })
            .catch((err) => {
                res.status(500).send(err);
            });
        }
        else
        {
            // error
            let errorMsg = "OAuth Access Token *not* found in session.\nPlease complete previous demo step\nto get an OAuth Access Token."; 
            Utils.logError(errorMsg);
            res.status(500).send(errorMsg);
        }
    }

    /**
     * loadDataHelper: uses the given OAuthAccessToklen to load JSON data into the Data Extension with external key "DF18Demo"
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.noversion.mc-apis.meta/mc-apis/postDataExtensionRowsetByKey.htm
     * 
     */
    private loadDataHelper(oauthAccessToken: string, jsonData: string) : Promise<any>    
    {
        let self = this;
        Utils.logInfo("loadDataHelper called.");
        Utils.logInfo("Loading sample data into Data Extension: " + self._deExternalKey);
        Utils.logInfo("Using OAuth token: " + this._oauthToken);

        return new Promise<any>((resolve, reject) =>
        {
            let headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + this._oauthToken
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios.post(self._sfmcDataExtensionApiUrl, jsonData, {"headers" : headers})
            .then((response: any) => {
                // success
                Utils.logInfo("Successfully loaded sample data into Data Extension!");

                resolve(
                {
                    status: response.status,
                    statusText: response.statusText + "\n" + Utils.prettyPrintJson(JSON.stringify(response.data))
                });
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error loading sample data. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
        });
    }
    
    public loadDataForPage2(req: express.Request, res: express.Response)
    {
        Utils.logInfo("request body = " + JSON.stringify(req.body));
        let self = this;
        let sessionId = req.session.id;
        Utils.logInfo("loadData entered. SessionId = " + sessionId);

        if (this._oauthToken!= "")
        {
            Utils.logInfo("Using OAuth token: " + this._oauthToken);
            self.loadDataHelperForPage2(this._oauthToken, JSON.stringify(req.body))
            .then((result) => {
                res.status(result.status).send(result.statusText);
            })
            .catch((err) => {
                res.status(500).send(err);
            });
        }
        else
        {
            // error
            let errorMsg = "OAuth Access Token *not* found in session.\nPlease complete previous demo step\nto get an OAuth Access Token."; 
            Utils.logError(errorMsg);
            res.status(500).send(errorMsg);
        }
    }

    /**
     * loadDataHelper: uses the given OAuthAccessToklen to load JSON data into the Data Extension with external key "DF18Demo"
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.noversion.mc-apis.meta/mc-apis/postDataExtensionRowsetByKey.htm
     * 
     */
    private loadDataHelperForPage2(oauthAccessToken: string, jsonData: string) : Promise<any>    
    {
        let self = this;
        Utils.logInfo("loadDataHelper called.");
        Utils.logInfo("Loading sample data into Data Extension: " + self._deExternalKey);
        Utils.logInfo("Using OAuth token: " + this._oauthToken);

        return new Promise<any>((resolve, reject) =>
        {
            let headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + oauthAccessToken
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios.post("https://mc4f63jqqhfc51yw6d1h0n1ns1.rest.marketingcloudapis.com/hub/v1/dataevents/key:" + "Data_Extension_Template" + "/rowset", jsonData, {"headers" : headers})
            .then((response: any) => {
                // success
                Utils.logInfo("Successfully loaded sample data into Data Extension!");

                resolve(
                {
                    status: response.status,
                    statusText: response.statusText + "\n" + Utils.prettyPrintJson(JSON.stringify(response.data))
                });
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error loading sample data. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
        });
		}
        
        public createDataExtension(req: express.Request, res: express.Response)
    {
	
        Utils.logInfo("request body for data extension creation = " + req.body.Template_Name);
        let self = this;
        let sessionId = req.session.id;
        Utils.logInfo("loadData entered. SessionId = " + sessionId);
	    //let template = req.body;

        if (this._oauthToken!= "")
        {
            //Utils.logInfo("Create Data extension method called and Condition satisfied: " + template );
            self.createDataExtensionHelper(this._oauthToken, req.body)
            .then((result) => {
                res.status(result.status).send(result.statusText);
            })
            .catch((err) => {
                res.status(500).send(err);
            });
        }
        else
        {
            // error
            let errorMsg = "OAuth Access Token *not* found in session.\nPlease complete previous demo step\nto get an OAuth Access Token."; 
            Utils.logError(errorMsg);
            res.status(500).send(errorMsg);
        }
    }

    
    private createDataExtensionHelper(oauthAccessToken: string, template: any) : Promise<any>    
    {
        let self = this;
		let bodySoapData = '';
		let sendableSoapData = '';
		let fieldSoapData = '';
		let endSoapData = '';
		let orgIDSoapData = '';
		let soapData = '';
		let templateNameData = '';
		let headerSoapData = '<?xml version="1.0" encoding="UTF-8"?>'
+' <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Create</a:Action>'
+'        <a:MessageID>urn:uuid:7e0cca04-57bd-4481-864c-6ea8039d2ea0</a:MessageID>'
+'        <a:ReplyTo>'
+'            <a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>'
+'        </a:ReplyTo>'
+'        <a:To s:mustUnderstand="1">https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+oauthAccessToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'        <CreateRequest xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'            <Options/>'





        //Utils.logInfo("createDataExtensionHelper method is called.");
        //Utils.logInfo("Using OAuth token: " + oauthAccessToken);
		//let dynamicTemplate = JSON.stringify(template);
		Utils.logInfo("Request body as a parameter: " + JSON.stringify(template));
		Object.keys(template).forEach(key => {
				Utils.logInfo(key);
				
				/*else{
					Utils.logInfo("Thambi, Hearsay User Reference ID innum varala");
				}*/
				if(key === "Template Name" ){
					Utils.logInfo("if condition satisfied");
					bodySoapData +=	'<Objects xsi:type="DataExtension">'
+'                <PartnerKey xsi:nil="true"/>'
+'                <ObjectID xsi:nil="true"/>'
+'                <CategoryID>'+this.FolderID+'</CategoryID>'
+'                <CustomerKey>'+template[key]+'</CustomerKey>'
+'                <Name>'+template[key]+'</Name>'
+'                <IsSendable>true</IsSendable>';

				templateNameData += '<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <MaxLength>50</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>';

		
				}
				else if(template[key]===""){
					Utils.logInfo("else if condition satisfied ");				
					delete template[key];
				}
				else if(key==="Hearsay Org ID"){
					//this.Hearsay_Org_ID = template[key];
					Utils.logInfo("Hearsay_Org_ID is blended with key, It will be sent as field name and the value inserted will be sent as value for that field ");				
					orgIDSoapData += '<Field>'
+'                        <Name>Org ID</Name>'
+'                        <DefaultValue>'+template[key]+'</DefaultValue>'
+'                        <MaxLength>50</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				else if(key === "Hearsay User Reference ID"){
					if(template[key]==="Email ID"){
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>EmailAddress</FieldType>'
+'                        <MaxLength>254</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
					}
					else{
						fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Text</FieldType>'
+'                        <MaxLength>100</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
					}

				}
				else if(template[key] ==="Email"){
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					sendableSoapData += '<SendableDataExtensionField>'
+'                    <CustomerKey>'+template[key]+'</CustomerKey>'
+'                    <Name>'+template[key]+'</Name>'
+'                    <FieldType>EmailAddress</FieldType>'
+'                </SendableDataExtensionField>'
+'                <SendableSubscriberField>'
+'                    <Name>Subscriber Key</Name>'
+'                    <Value>'+template[key]+'</Value>'
+'                </SendableSubscriberField>'
+'                <Fields>'
+'					<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>EmailAddress</FieldType>'
+'                        <MaxLength>254</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				else if (key==="option 2" || key==="option 3" || key==="option 4" ||key==="option 5" ||key==="option 6" ||key==="option 7" ||key==="option 8" ||key==="option 9" ){
				if(template[key] ==="Birth Date"){
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Date</FieldType>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				else if(template[key] ==="Phone"){
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Phone</FieldType>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				else{
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Phone</FieldType>'
+'                        <IsRequired>false</IsRequired>'
+'                    </Field>'
				}
				}
				else{
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					fieldSoapData += '<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Text</FieldType>'
+'                        <MaxLength>50</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				
			});
			endSoapData += '</Fields>'
+'            </Objects>'
+'        </CreateRequest>'
+'    </s:Body>'
+'</s:Envelope>';
			
			Utils.logInfo("Request body after deletion: " + JSON.stringify(template));
			
 soapData = headerSoapData + bodySoapData + sendableSoapData + templateNameData + orgIDSoapData + fieldSoapData + endSoapData;	    
	    
        return new Promise<any>((resolve, reject) =>
        {
				
				axios({
				method: 'post',
				url: 'https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx',
				data: soapData,
				headers: {'Content-Type': 'text/xml'}							
				})            
				.then((response: any) => {
                Utils.logInfo(response.data);
                var extractedData = "";
				/*var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				//this.FolderID = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ID'][0];
				//Utils.logInfo('Folder ID : ' + this.FolderID);
				this.StatusCode = result['soap:Envelope']['soap:Body'][0]['CreateResponse'][0]['Results'][0]['StatusCode'];
				Utils.logInfo('Status Code : ' + this.StatusCode);
				if(this.StatusCode=='OK'){
					this.DataExtensionName = result['soap:Envelope']['soap:Body'][0]['CreateResponse'][0]['Results'][0]['Object'][0]['Name'];
					Utils.logInfo('Data Extension Name : ' + this.DataExtensionName);
					
					self.RowCreationDynamicDataExt(this.DataExtensionName)
					.then((result : any) => {
						Utils.logInfo('Row Created Successfully in Dynamically created DE');
						})
					.catch((err : any) => {
						Utils.logInfo('Error when creating row in Dynamically created DE');
					});
				}
				else{
					Utils.logInfo('Olunga odi poiru condition ae satisfy aagala');
				}
				});*/
				})
		/*let headers = {
                'Content-Type': 'text/xml'
            };
			
            axios.post('https://mc4f63jqqhfc51yw6d1h0n1ns1.soap.marketingcloudapis.com/Service.asmx', soapData, {"headers" : headers})
			.then(function (response) {
				Utils.logInfo(response.data);
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				//this.FolderID = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ID'][0];
				//Utils.logInfo('Folder ID : ' + this.FolderID);
				this.StatusCode = result['soap:Envelope']['soap:Body'][0]['CreateResponse'][0]['Results'][0]['StatusCode'];
				Utils.logInfo('Status Code : ' + this.StatusCode);
				});
			})*/
			.catch(function (error) {
				let errorMsg = "Error creating data extension dynamically. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);
				Utils.logError(error.response.data);

                reject(errorMsg);
			});
        });
    }
	
	/*private RowCreationDynamicDataExt(DataExtensionName : any) : Promise<any>
	{
		
		let RowData=
	    {
			items: [{
			Org_ID : this.Hearsay_Org_ID
			}]                                		
	    }
			let Row = JSON.stringify(RowData);
			Utils.logInfo("Row "+ Row);
			
			let headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + this._oauthToken
            };
		
		return new Promise<any>((resolve, reject) =>
        {
			Utils.logInfo("Ahpppaaaddaa, Method call aaiduchu");
			 axios.post("https://mc4f63jqqhfc51yw6d1h0n1ns1.rest.marketingcloudapis.com/data/v1/async/dataextensions/key:" + DataExtensionName + "/rows", Row, {"headers" : headers})
            .then((response: any) => {
                // success
                Utils.logInfo("Hearsay_Org_ID Updated Successfully");

                
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error loading sample data. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
			
		});
	}*/
        
    }

