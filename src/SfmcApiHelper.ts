'use strict';

import axios from 'axios';
import express = require("express");
import Utils from './Utils';

export default class SfmcApiHelper
{
    // Instance variables 
    private _deExternalKey = "OrgSetup";
    private _sfmcDataExtensionApiUrl = "https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.rest.marketingcloudapis.com/hub/v1/dataevents/key:" + this._deExternalKey + "/rowset";
    
    
    
    
    
    
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
            let sfmcAuthServiceApiUrl = "https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.auth.marketingcloudapis.com/v2/token";
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
                Utils.logInfo("tokenExpiry..." + tokenExpiry);
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

        if (req.session.oauthAccessToken)
        {
            Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.loadDataHelper(req.session.oauthAccessToken, JSON.stringify(req.body))
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
        Utils.logInfo("Using OAuth token: " + oauthAccessToken);

        return new Promise<any>((resolve, reject) =>
        {
            let headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + oauthAccessToken
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

        if (req.session.oauthAccessToken)
        {
            Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.loadDataHelperForPage2(req.session.oauthAccessToken, JSON.stringify(req.body))
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
        Utils.logInfo("Using OAuth token: " + oauthAccessToken);

        return new Promise<any>((resolve, reject) =>
        {
            let headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + oauthAccessToken
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios.post("https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.rest.marketingcloudapis.com/hub/v1/dataevents/key:" + "page2" + "/rowset", jsonData, {"headers" : headers})
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
        Utils.logInfo("request body = " + req.body);
        let self = this;
        let sessionId = req.session.id;
        Utils.logInfo("loadData entered. SessionId = " + sessionId);

        if (req.session.oauthAccessToken)
        {
            Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.createDataExtensionHelper(req.session.oauthAccessToken, req.body)
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

    
    private createDataExtensionHelper(oauthAccessToken: string, jsonData: string) : Promise<any>    
    {
        let self = this;
        Utils.logInfo("createDataExtensionHelper method is called.");
        //Utils.logInfo("Loading sample data into Data Extension: " + self._deExternalKey);
        Utils.logInfo("Using OAuth token: " + oauthAccessToken);
	    
	    var data = '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
 +'\n  <soap:Header>'
+'\n      <wsa:Action>Create</wsa:Action>'
+'\n      <wsa:MessageID>urn:uuid:168bbf3d-394e-4656-ae57-2e96b4b568ae</wsa:MessageID>'
+'\n     <wsa:ReplyTo>'
+'\n         <wsa:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</wsa:Address>'
+'\n      </wsa:ReplyTo>'
+'\n      <wsa:To>https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.soap.marketingcloudapis.com/Service.asmx</wsa:To>'
+'\n      <wsse:Security soap:mustUnderstand="1">'
+'\n         <wsse:UsernameToken wsu:Id="SecurityToken-d19fb7b0-ec6d-49a8-8fd3-796819ec7306">'
+'\n            <wsse:Username>sivaorkestra</wsse:Username>'
+'\n            <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText">pashtekjuly2120#!</wsse:Password>'
+'\n         </wsse:UsernameToken>'
+'\n      </wsse:Security>'
+'\n   </soap:Header>'
+'\n   <soap:Body>'
+'\n      <CreateRequest xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'\n         <Options>'
+'\n            <SaveOptions/>'
+'\n         </Options>'
+'\n         <Objects xsi:type="DataExtension">'
+'\n            <PartnerKey xsi:nil="true"/>'
+'\n            <ObjectID xsi:nil="true"/>'
+'\n            <CustomerKey>02_2010-02-26-08_42_30_762_851628611</CustomerKey>'
+'\n            <Name>02_2010-02-26-08_42_30_762_851628611</Name>'
+'\n            <Description>02_2010-02-26-08_42_30_762_851628611</Description>'
+'\n            <IsSendable>true</IsSendable>'
+'\n            <IsTestable>false</IsTestable>'
+'\n            <DataRetentionPeriodLength>48</DataRetentionPeriodLength>'
+'\n            <DataRetentionPeriodUnitOfMeasure>0</DataRetentionPeriodUnitOfMeasure>'
+'\n            <RowBasedRetention>false</RowBasedRetention>'
+'\n            <ResetRetentionPeriodOnImport>true</ResetRetentionPeriodOnImport>'
+'\n            <DeleteAtEndOfRetentionPeriod>false</DeleteAtEndOfRetentionPeriod>'
 +'\n           <Fields>'
+'\n               <Field>'
+'\n                  <PartnerKey xsi:nil="true"/>'
+'\n                  <ObjectID xsi:nil="true"/>'
+'\n                  <Name>D_Body</Name>'
+'\n                  <Description>D_Body</Description>'
+'\n                  <IsRequired>true</IsRequired>'
 +'\n                 <IsPrimaryKey>false</IsPrimaryKey>'
+'\n                  <FieldType>Text</FieldType>'
+'\n                  <DefaultValue></DefaultValue>'
+'\n               </Field>'
+'\n               <Field>'
+'\n                  <PartnerKey xsi:nil="true"/>'
+'\n                  <ObjectID xsi:nil="true"/>'
+'\n                  <Name>D_F1</Name>'
+'\n                  <Description>D_F1</Description>'
+'\n                  <MaxLength>100</MaxLength>'
 +'\n                 <IsRequired>true</IsRequired>'
+'\n                  <IsPrimaryKey>false</IsPrimaryKey>'
+'\n                  <FieldType>Text</FieldType>'
+'\n                  <DefaultValue></DefaultValue>'
 +'\n              </Field>'
+'\n               <Field>'
 +'\n                 <PartnerKey xsi:nil="true"/>'
+'\n                  <ObjectID xsi:nil="true"/>'
+'\n                  <Name>D_F2</Name>'
+'\n                  <Description>D_F2</Description>'
+'\n                  <MaxLength>100</MaxLength>'
+'\n                  <IsRequired>true</IsRequired>'
+'\n                  <IsPrimaryKey>false</IsPrimaryKey>'
+'\n                  <FieldType>Text</FieldType>'
+'\n                  <DefaultValue></DefaultValue>'
+'\n               </Field>'               
+'\n            </Fields>'
+'\n            <SendableDataExtensionField>'
+'\n               <PartnerKey xsi:nil="true"/>'
+'\n               <ObjectID xsi:nil="true"/>'
+'\n               <Name>Email Address</Name>'
+'\n            </SendableDataExtensionField>'
+'\n            <SendableSubscriberField>'
+'\n               <Name>Email Address</Name>'
+'\n            </SendableSubscriberField>'
+'\n         </Objects>'
+'\n      </CreateRequest>'
+'\n   </soap:Body>'
+'\n </soap:Envelope>';

        return new Promise<any>((resolve, reject) =>
        {
			let headers = {
                'Content-Type': 'text/xml'
            };
			
            axios.post('https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.soap.marketingcloudapis.com/Service.asmx', data, {"headers" : headers})
			.then(function (response) {
				Utils.logInfo(response.data);
			})
			.catch(function (error) {
				let errorMsg = "Error creating data extension dynamically. POST response from Marketing Cloud:";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
			});
        });
    }
        
    }

