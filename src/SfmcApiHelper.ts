'use strict';

import axios from 'axios';
import express = require("express");
import Utils from './Utils';
import xml2js = require("xml2js");

export default class SfmcApiHelper
{
    // Instance variables 
    private _deExternalKey = "Org_Setup";
    
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
	private ValidateResponse = '';
	private soap_instance_url ='';
	private rest_instance_url = '';
	private member_id = '';
    
    
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
            let sfmcAuthServiceApiUrl = process.env.BASE_URL +"auth.marketingcloudapis.com/v2/token";
            Utils.logInfo("oauth token is called, waiting for status...");
            axios.post(sfmcAuthServiceApiUrl, postBody, {"headers" : headers})            
            .then((response : any) => {
                // success
                Utils.logInfo("Success, got auth token from MC..." + Utils.prettyPrintJson(JSON.stringify(response.data)));
                let accessToken = response.data.access_token;
                Utils.logInfo("oauth token..." + accessToken);
                let bearer = response.data.token_type;
                Utils.logInfo("Bearer..." + bearer);
                let tokenExpiry = response.data.expires_in;
                Utils.logInfo("tokenExpiry..." + tokenExpiry);
                
				this._oauthToken = response.data.access_token;
				Utils.logInfo("Storing the accesstoken in a object's variable "+ this._oauthToken+"\n");
				

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
	
	public DataExtensionFolderCheck(req: express.Request, res: express.Response){
		Utils.logInfo("DataExtensionFolderCheck Method: " + this._oauthToken);
		res.status(200).send("The The The Tha Tha Tha The The The");
		this.getCategoryIDHelper()
		.then((result) => {
                res.status(result.status).send(result.statusText);
            })
            .catch((err) => {
                res.status(500).send(err);
            });
		
	}
	
	public AppUserInfo(req: express.Request, res: express.Response){
		Utils.logInfo("App User Information method called " + this._oauthToken);
		let self = this;
		
		let userInfoUrl = process.env.BASE_URL +"auth.marketingcloudapis.com/v2/userinfo";
		
		let headers = {
            'Content-Type': 'application/json',
			'Authorization': 'Bearer ' + this._oauthToken
        };
		
		axios.get(userInfoUrl,{"headers" : headers})            
            .then((response : any) => {
                // success
                Utils.logInfo("User Information response Body : "+ JSON.stringify(response.data));
				Utils.logInfo("Member ID of the Current Business Unit : "+ Utils.prettyPrintJson(JSON.stringify(response.data.organization.member_id)));
				this.member_id = response.data.organization.member_id;
				this.soap_instance_url = response.data.rest.soap_instance_url;
				Utils.logInfo("Storing the Soap URL in a object's variable "+ this.soap_instance_url+"\n");
                this.rest_instance_url = response.data.rest.rest_instance_url;
				Utils.logInfo("Storing the Rest URL in a object's variable "+ this.rest_instance_url+"\n");
				res.status(200).send(Utils.prettyPrintJson(JSON.stringify(response.data)));
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error getting User's Information.";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                res.status(500).send(Utils.prettyPrintJson(JSON.stringify(error.response.data)));
            });
		//res.status(200).send("AppUserInfo Method Successful a call aaiduchu da perapasangala");
		
	}
	
	public getCategoryID(req: express.Request, res: express.Response)
    {
		Utils.logInfo("Get Category Method: " + this._oauthToken);
		let self = this;	
			
		Utils.logInfo("request body = " + JSON.stringify(req.query.TemplateName));
		let TemplateName = req.query.TemplateName;
		if (this._oauthToken!= "")
        {
            //Utils.logInfo("Using OAuth token: " + req.session.oauthAccessToken);
            self.ValidationForDataExtName(TemplateName)
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
	
	private ValidationForDataExtName(TemplateName : string) : Promise<any>
	{
		
			//Utils.logInfo("Validation Body : "+ ValidationBody);
			
			//let headers = {
                //'Content-Type': 'application/json',
                //
            //};
		
		return new Promise<any>((resolve, reject) =>
        {
			
			let validateName = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Retrieve</a:Action>'
+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'   <RetrieveRequestMsg xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'      <RetrieveRequest>'
+'		<ObjectType>DataExtension</ObjectType>'
+'                <Properties>CustomerKey</Properties>'
+'                <Properties>Name</Properties>'
+'                <Filter xsi:type="SimpleFilterPart">'
+'                  <Property>Name</Property>'
+'                  <SimpleOperator>equals</SimpleOperator>'
+'                  <Value>'+TemplateName+'</Value>'
+'                </Filter>'
+'      </RetrieveRequest>'
+'   </RetrieveRequestMsg>'
+'    </s:Body>'
+'</s:Envelope>';
			Utils.logInfo("Ahpppaaaddaa, Method call aaiduchu");
			 axios({
				method: 'post',
				url: ''+this.soap_instance_url+'Service.asmx'+'',
				data: validateName,
				headers: {'Content-Type': 'text/xml'}							
				}) 
            .then((response: any) => {
                // success
                Utils.logInfo("Validation Successful \n\n" + response.data);
				this.ValidateResponse = response.data;
				
				let parser = new xml2js.Parser();
				let parsedResponse = parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				//this.validateStatus = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['OverallStatus'][0];
				//Utils.logInfo('Validation Status : ' + response.data);
				
				
				let validateDEName = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'];
				Utils.logInfo('Validated Data Extension Name : ' + validateDEName);
				if(validateDEName!=undefined){
				resolve(
                {
                    status: 200,
                    statusText: validateDEName[0]['Name'][0]
                });
				}
				else{
					validateDEName = "null";
					resolve({
                    status: 200,
                    statusText: validateDEName
                });
				}
				
                /*if(validateDEName){
					Utils.logInfo('Validated Data Extension Name : ' + validateDEName);
				resolve(
                {
                    status: 200,
                    statusText: validateDEName[0]['Properties'][0]['Property'][0]['Value'][0]
                });
			}*/
				
			
				});
				
			
            })
            .catch((error: any) => {
                // error
                let errorMsg = "Error retrieving the data extension name :";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

                reject(errorMsg);
            });
			
	})
	}
	
	
	public getCategoryIDHelper() : Promise<any>
	{
		Utils.logInfo('getCategoryIDHelper Method has been called, Finally it Worked. We are going to complete');
		let soapMessage = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Retrieve</a:Action>'
+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
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
				url: ''+this.soap_instance_url+'Service.asmx'+'',
				data: soapMessage,
				headers: {'Content-Type': 'text/xml'}							
				})            
				.then((response: any) => {
                Utils.logInfo(response.data);
                var extractedData = "";
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				let FolderID = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'];
				Utils.logInfo('Folder ID : ' + FolderID);
				//this.ParentFolderID = JSON.stringify(result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ParentFolder'][0]);
				//Utils.logInfo('Parent Folder ID : ' + this.ParentFolderID);
					if(FolderID!=undefined){
				this.FolderID = FolderID[0]['ID'][0];
				}
				else{
					this.retrievingDataExtensionFolderID();
					
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
			
        });
	}
	
	public retrievingDataExtensionFolderID(){
		
		Utils.logInfo('Retrieving DataExtension Folder Properties......');
		let soapMessage = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Retrieve</a:Action>'
+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'        <RetrieveRequestMsg xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'            <RetrieveRequest>'
+'                <ObjectType>DataFolder</ObjectType>'
+'                <Properties>ID</Properties>'
+'                <Properties>CustomerKey</Properties>'
+'                <Properties>Name</Properties>'
+'                <Filter xsi:type="SimpleFilterPart">'
+'                    <Property>Name</Property>'
+'                    <SimpleOperator>equals</SimpleOperator>'
+'                    <Value>Data Extensions</Value>'
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
				url: ''+this.soap_instance_url+'Service.asmx'+'',
				data: soapMessage,
				headers: {'Content-Type': 'text/xml'}							
				})            
				.then((response: any) => {
                Utils.logInfo(response.data);
                var extractedData = "";
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				let ParentFolderID = result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ID'][0];
				Utils.logInfo('Idhu dhan da parentFolder ID en vendru : ' + ParentFolderID);
				//this.ParentFolderID = JSON.stringify(result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ParentFolder'][0]);
				//Utils.logInfo('Parent Folder ID : ' + this.ParentFolderID);
					
						if(ParentFolderID!=undefined){
							this.ParentFolderID = ParentFolderID;
							this.creatingHearsayIntegrationFolder(ParentFolderID);
							
						}
				
				});
				})
			.catch((error: any) => {
						// error
						let errorMsg = "Error getting the Data extensions folder properties......";
						errorMsg += "\nMessage: " + error.message;
						errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
						errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
						Utils.logError(errorMsg);

										reject(errorMsg);
									});
			
        });
}
	
	public creatingHearsayIntegrationFolder(ParentFolderID : any){
		
		let createFolderData = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                                 +'<soapenv:Header>'
                                 +'<fueloauth>'+this._oauthToken+'</fueloauth>'
                                 +'</soapenv:Header>'
                                 +'<soapenv:Body>'
                                 +'<CreateRequest xmlns="http://exacttarget.com/wsdl/partnerAPI">'
                                 +'<Options/>'
                                 +'<ns1:Objects xmlns:ns1="http://exacttarget.com/wsdl/partnerAPI" xsi:type="ns1:DataFolder">'
                                 +'<ns1:ModifiedDate xsi:nil="true"/>'
                                 +'<ns1:ObjectID xsi:nil="true"/>'
                                 +'<ns1:CustomerKey>Hearsay Integrations</ns1:CustomerKey>'
                                 +'<ns1:ParentFolder>'
                                 +'<ns1:ModifiedDate xsi:nil="true"/>'
                                 +'<ns1:ID>'+ParentFolderID+'</ns1:ID>'
                                 +'<ns1:ObjectID xsi:nil="true"/>'
                                 +'</ns1:ParentFolder>'
                                 +'<ns1:Name>Hearsay Integrations</ns1:Name>'
                                 +'<ns1:Description>Hearsay Integrations Folder</ns1:Description>'
                                 +'<ns1:ContentType>dataextension</ns1:ContentType>'
                                 +'<ns1:IsActive>true</ns1:IsActive>'
                                 +'<ns1:IsEditable>true</ns1:IsEditable>'
                                 +'<ns1:AllowChildren>true</ns1:AllowChildren>'
                                 +'</ns1:Objects>'
                                 +'</CreateRequest>'
                                 +'</soapenv:Body>'
                                 +'</soapenv:Envelope>';
		
		return new Promise<any>((resolve, reject) =>
		{
			let headers = {
                'Content-Type': 'text/xml',
                'SOAPAction': 'Create'
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios({
				method: 'post',
				url: ''+this.soap_instance_url+'Service.asmx'+'',
				data: createFolderData,
				headers: headers							
				})            
				.then((response: any) => {
					
					Utils.logInfo("Hearsay Integrations Folder has been created Successfully");
                Utils.logInfo(response.data);
				
				//Utils.logInfo("Folder ID nu onnu naa create pannae adhoda value a check panrae : " +this.FolderID);
				//this.createDefaultDataExtension
                var extractedData = "";
				var parser = new xml2js.Parser();
				parser.parseString(response.data, (err: any, result: { [x: string]: { [x: string]: { [x: string]: { [x: string]: any; }[]; }[]; }; }) => {
				let HearsayIntegrationsID = result['soap:Envelope']['soap:Body'][0]['CreateResponse'][0]['Results'][0]['NewID'][0];
				Utils.logInfo('Hearsay Integrations Folder ID : ' + HearsayIntegrationsID);
				//this.ParentFolderID = JSON.stringify(result['soap:Envelope']['soap:Body'][0]['RetrieveResponseMsg'][0]['Results'][0]['ParentFolder'][0]);
				//Utils.logInfo('Parent Folder ID : ' + this.ParentFolderID);
					
						if(HearsayIntegrationsID!=undefined){
							this.FolderID = HearsayIntegrationsID;
							this.creatingDefaultDataExtensions();							
						}
				
				});
				})
			.catch((error: any) => {
						// error
						let errorMsg = "Error creating the Hearsay Integrations folder......";
						errorMsg += "\nMessage: " + error.message;
						errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
						errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
						Utils.logError(errorMsg);

										reject(errorMsg);
									});
			
        });
	}
	
	public creatingDefaultDataExtensions(){
		
		Utils.logInfo("Creating Default Data Extensions for Org Setup");
		
		/*let OrgMsg = '<?xml version="1.0" encoding="UTF-8"?>'
		+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
		+'    <s:Header>'
		+'        <<a:Action s:mustUnderstand="1">Create</a:Action>'
		+'        <a:MessageID>urn:uuid:7e0cca04-57bd-4481-864c-6ea8039d2ea0</a:MessageID>'
		+'        <a:ReplyTo>'
		+'            <a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>'
		+'        </a:ReplyTo>'
		+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
		+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
		+'    </s:Header>'
		+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
		+'        <CreateRequest xmlns="http://exacttarget.com/wsdl/partnerAPI">'
		+'<Objects xsi:type="DataExtension">'
		+'<CategoryID>'+this.FolderID+'</CategoryID>'
		+'<CustomerKey>Org_Setup</CustomerKey>'
        	+'<Name>Org Setup</Name>'
        	+'<IsSendable>true</IsSendable>'
        	+'<SendableDataExtensionField>'
		+'    <CustomerKey>Hearsay User Reference ID</CustomerKey>'
		+'    <Name>Hearsay User Reference ID</Name>'
		+'    <FieldType>Text</FieldType>'
		+'</SendableDataExtensionField>'
        	+'<SendableSubscriberField>'
        	+'    <Name>Subscriber Key</Name>'
        	+'    <Value></Value>'
        	+'</SendableSubscriberField>'
		+'<Fields>'
		+'<CustomerKey>Hearsay Org ID</CustomerKey>'
		+'<Name>Hearsay Org ID</Name>'
		+'<FieldType>Text</FieldType>'
		+'<MaxLength>50</MaxLength>'
		+'<IsRequired>true</IsRequired>'
		+'<IsPrimaryKey>false</IsPrimaryKey>'
		+'</Field>'
	    	+'<Field>'
		+'<CustomerKey>Hearsay User Reference ID</CustomerKey>'
		+'<Name>Hearsay User Reference ID</Name>'
		+'<FieldType>Text</FieldType>'
		+'<MaxLength>50</MaxLength>'
		+'<IsRequired>true</IsRequired>'
		+'<IsPrimaryKey>true</IsPrimaryKey>'
		+'</Field>'
	    	+'<Field>'
		+'<CustomerKey>Customer Unique ID</CustomerKey>'
		+'<Name>Customer Unique ID</Name>'
		+'<FieldType>Text</FieldType>'
		+'<MaxLength>50</MaxLength>'
		+'<IsRequired>true</IsRequired>'
		+'<IsPrimaryKey>false</IsPrimaryKey>'
		+'</Field>'
	        +'<Field>'
		+'<CustomerKey>Created or Modified Date</CustomerKey>'
		+'<Name>Created or Modified Date</Name>'
		+'<FieldType>Date</FieldType>'
	        +'<DefaultValue>GetDate()</DefaultValue>'
		+'<IsRequired>true</IsRequired>'
		+'<IsPrimaryKey>false</IsPrimaryKey>'
		+'</Field>'
	    	+'</Fields>'
		+'</Objects>'
		+'        </CreateRequest>'
		+'    </s:Body>'
		+'</s:Envelope>';*/
		
		let OrgMsg = '<?xml version="1.0" encoding="UTF-8"?>'
+'<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Create</a:Action>'
+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
+'        <fueloauth xmlns="http://exacttarget.com">'+this._oauthToken+'</fueloauth>'
+'    </s:Header>'
+'    <s:Body xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
+'        <CreateRequest xmlns="http://exacttarget.com/wsdl/partnerAPI">'
+'            <Objects xsi:type="DataExtension">'
+'				<CategoryID>'+this.FolderID+'</CategoryID>'
+'                <Client>'
+'                    <ID>'+this.member_id+'</ID>'
+'                </Client>'
+'                <CustomerKey>Org_Setup</CustomerKey>'
+'                <Name>Org Setup</Name>'
+'                <IsSendable>true</IsSendable>'
+'                <SendableDataExtensionField>'
+'                    <CustomerKey>Customer Unique ID</CustomerKey>'
+'                    <Name>Customer Unique ID</Name>'
+'                    <FieldType>EmailAddress</FieldType>'
+'                </SendableDataExtensionField>'
+'                <SendableSubscriberField>'
+'                    <Name>Subscriber Key</Name>'
+'                    <Value></Value>'
+'                </SendableSubscriberField>'
+'                <Fields>'
+'                    <Field>'
+'                        <CustomerKey>Customer Unique ID</CustomerKey>'
+'                        <Name>Customer Unique ID</Name>'
+'                        <FieldType>EmailAddress</FieldType>'
+'                        <MaxLength>254</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                        <IsPrimaryKey>true</IsPrimaryKey>'
+'                    </Field>'
+'                    <Field>'
+'                        <CustomerKey>Hearsay Org ID</CustomerKey>'
+'                        <Name>Hearsay Org ID</Name>'
+'                       <FieldType>Text</FieldType>'
+'                        <MaxLength>50</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                        <IsPrimaryKey>false</IsPrimaryKey>'
+'                    </Field>'
+'                    <Field>'
+'                        <CustomerKey>Hearsay User Reference ID</CustomerKey>'
+'                        <Name>Hearsay User Reference ID</Name>'
+'                        <FieldType>Text</FieldType>'
+'                        <MaxLength>50</MaxLength>'
+'                        <IsRequired>true</IsRequired>'
+'                        <IsPrimaryKey>false</IsPrimaryKey>'
+'                    </Field>'
+'                    <Field>'
+'                        <CustomerKey>Created or Modified Date</CustomerKey>'
+'                        <Name>Created or Modified Date</Name>'
+'                        <FieldType>Date</FieldType>'
+'						<DefaultValue>getdate()</DefaultValue>'
+'                       <IsRequired>true</IsRequired>'
+'                        <IsPrimaryKey>false</IsPrimaryKey>'
+'                    </Field>'
+'                </Fields>'
+'            </Objects>'
+'        </CreateRequest>'
+'    </s:Body>'
+'</s:Envelope>'
		
		
		
		
		
		Utils.logInfo("The soap data for creating the Org Setup Data Extension\n\n\n" + OrgMsg);
		
		return new Promise<any>((resolve, reject) =>
		{
			let headers = {
                'Content-Type': 'text/xml'
            };

            // POST to Marketing Cloud Data Extension endpoint to load sample data in the POST body
            axios({
				method: 'post',
				url: ''+this.soap_instance_url+'Service.asmx'+'',
				data: OrgMsg,
				headers: headers							
				})            
				.then((response: any) => {
					
				Utils.logInfo("Org Setup Data extension has been created Successfully\n\n\n");
                Utils.logInfo(response.data+"\n\n\n");				
				Utils.logInfo("Oru valiya idhuvum create panni aachu, innum Data Extension Template mattum dha mitcham\n\n\n");					
				})
			.catch((error: any) => {
						// error
						let errorMsg = "Error creating the Org Setup Data extension......";
						errorMsg += "\nMessage: " + error.message;
						errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
						errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
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
		let _sfmcDataExtensionApiUrl = this.rest_instance_url + "/hub/v1/dataevents/key:" + this._deExternalKey + "/rowset";
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
            axios.post(_sfmcDataExtensionApiUrl, jsonData, {"headers" : headers})
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
            axios.post(this.rest_instance_url+"/hub/v1/dataevents/key:" + "Data_Extension_Template" + "/rowset", jsonData, {"headers" : headers})
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
		let OptionFieldSoapData = '';
		let headerSoapData = '<?xml version="1.0" encoding="UTF-8"?>'
+' <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">'
+'    <s:Header>'
+'        <a:Action s:mustUnderstand="1">Create</a:Action>'
+'        <a:MessageID>urn:uuid:7e0cca04-57bd-4481-864c-6ea8039d2ea0</a:MessageID>'
+'        <a:ReplyTo>'
+'            <a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address>'
+'        </a:ReplyTo>'
+'        <a:To s:mustUnderstand="1">'+this.soap_instance_url+'Service.asmx'+'</a:To>'
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
					
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>EmailAddress</FieldType>'
+'                        <MaxLength>254</MaxLength>'
+'                        <IsRequired>false</IsRequired>'
+'                    </Field>'	
					
				}
				else if(template[key] ==="Birth Date"){
					Utils.logInfo("field name Birth Date of "+key+"has been added to the soapData");
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Date</FieldType>'
+'                        <IsRequired>false</IsRequired>'
+'                    </Field>'
				}
				else if(key ==="option 2"){
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Phone</FieldType>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'

				}
				else if(template[key] ==="Name"){
					Utils.logInfo("field name "+ template[key] + " has been added to the soapData");
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Text</FieldType>'
+'                        <IsRequired>true</IsRequired>'
+'                    </Field>'
				}
				else if(key ==="Customer Unique ID"){
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
				
				
				
				else if(key ==="option 9" || key ==="option 10" || key ==="option 11"){
					Utils.logInfo("field name "+ template[key] + " has been added to the OptionFieldSoapData");
					OptionFieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'

				}
				else if(key ==="FieldType 1" || key ==="FieldType 2" || key ==="FieldType 3"){
					if(template[key] === "Date"){
						OptionFieldSoapData +='<FieldType>'+template[key]+'</FieldType>'
					+'                        <IsRequired>false</IsRequired>'
					+'                    </Field>'
					}
					else{
					Utils.logInfo("field type "+ template[key] + " has been added to the OptionFieldSoapData");
					OptionFieldSoapData +='<FieldType>'+template[key]+'</FieldType>'
					}
				}
				
				/*else if(key ==="FieldType 1" && template[key] === "Date"|| key ==="FieldType 2" && template[key] === "Date"|| key ==="FieldType 3" && template[key] === "Date"){
					Utils.logInfo("This is to check whether the selected data type is Date");
					OptionFieldSoapData +='<FieldType>'+template[key]+'</FieldType>'
					+'                        <IsRequired>false</IsRequired>'
					+'                    </Field>'
				}*/
				
				else if(key ==="FieldLength 1" || key ==="FieldLength 2" || key ==="FieldLength 3"){
					Utils.logInfo("field length "+ template[key] + " has been added to the OptionFieldSoapData");
					OptionFieldSoapData += '<MaxLength>'+ template[key] +'</MaxLength>'
+'                        <IsRequired>false</IsRequired>'					
+'                    </Field>'
				}
				
				
				else{
					fieldSoapData +='<Field>'
+'                        <Name>'+template[key]+'</Name>'
+'                        <FieldType>Text</FieldType>'
+'                        <IsRequired>false</IsRequired>'
+'                    </Field>'
				}				
			});
			endSoapData += '</Fields>'
+'            </Objects>'
+'        </CreateRequest>'
+'    </s:Body>'
+'</s:Envelope>';
			
			Utils.logInfo("Request body after deletion: " + JSON.stringify(template));
			
 soapData = headerSoapData + bodySoapData + sendableSoapData + orgIDSoapData + fieldSoapData + OptionFieldSoapData + endSoapData;	
 
	    Utils.logInfo("Soap data before the data extension creation call\n " + soapData);
        return new Promise<any>((resolve, reject) =>
        {
				
				axios({
				method: 'post',
				url: this.soap_instance_url+'Service.asmx',
				data: soapData,
				headers: {'Content-Type': 'text/xml'}							
				})            
				.then((response: any) => {
                Utils.logInfo(response.data);
                //var extractedData = "";
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
				
				resolve(
                {
                    status: response.status,
                    statusText: response.statusText + "\n" + Utils.prettyPrintJson(JSON.stringify(response.data))
                });
				})
		/*let headers = {
                'Content-Type': 'text/xml'
            };
			
            axios.post('https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.soap.marketingcloudapis.com/Service.asmx', soapData, {"headers" : headers})
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
			.catch((error: any) => {
                // error
                let errorMsg = "Error Creating data extension dynamically using soap format";
                errorMsg += "\nMessage: " + error.message;
                errorMsg += "\nStatus: " + error.response ? error.response.status : "<None>";
                errorMsg += "\nResponse data: " + error.response.data ? Utils.prettyPrintJson(JSON.stringify(error.response.data)) : "<None>";
                Utils.logError(errorMsg);

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
			 axios.post("https://mcj6cy1x9m-t5h5tz0bfsyqj38ky.rest.marketingcloudapis.com/data/v1/async/dataextensions/key:" + DataExtensionName + "/rows", Row, {"headers" : headers})
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
	
	/*public logout() 
	{
                response.redirect('https://mc.s11.exacttarget.com/');
				})
    }*/
}
