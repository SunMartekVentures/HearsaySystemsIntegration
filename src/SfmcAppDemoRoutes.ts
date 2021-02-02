'use strict';

import express = require("express");
import jwt = require('jwt-simple');
import SfmcApiHelper from './SfmcApiHelper';
import Utils from './Utils';


// <!-- Integrate an externally hosted app via iframe. -->
export default class SfmcAppDemoRoutes
{
    // Instance variables
    private _apiHelper = new SfmcApiHelper();

    /**
     * login: called by Marketing Cloud when hosted app is launched. Decodes JWT in BODY passed by Marketing Cloud.
     * Handles POST on: /login
     * 
     * Marketing Cloud does a POST on the '/login' route with the following JSON BODY:
     * {
     *  "jwt" : "<encoded JWT from SFMC>"
     * }
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.mc-app-development.meta/mc-app-development/decode-jwt.htm
     * More info: https://developer.salesforce.com/docs/atlas.en-us.mc-app-development.meta/mc-app-development/explanation-decoded-jwt.htm
     * 
     */
    public login(req: express.Request, res: express.Response)
    {
        let self = this;
        let sessionId = req.session.id;
        Utils.logInfo("login called. SessionId = " + sessionId);

        req.session.jwtFromSFMC = "";
        req.session.decodedJWT = "";
        req.session.refreshTokenFromJWT = "";
        req.session.oauthAccessToken = "";
        req.session.oauthAccessTokenExpiry = "";

        // Decode JWT with the secret from environment variable.ggg
        try
        {
            // Decode JWT
            let encodedJWT = JSON.stringify(req.body.jwt);
            let jwtSecret = process.env.DF18DEMO_JWTSECRET;
            Utils.logInfo("Decoding JWT with secret from DF18DEMO_JWTSECRET = " + jwtSecret);
            req.session.decodedJWT = jwt.decode(encodedJWT, jwtSecret, true); // pass 'noVerify = true' for this demo
            
            // Store JWT in the current session for redirect URL to pick up for display
            req.session.jwtFromSFMC = Utils.prettyPrintJson(JSON.stringify(req.session.decodedJWT));
            Utils.logInfo("Decoded JWT from SFMC = \n" + req.session.jwtFromSFMC);

            // Get refreshToken from JWT and store in the current session for redirect URL to pick up for display
            req.session.refreshTokenFromJWT = req.session.decodedJWT.request.rest.refreshToken;
            Utils.logInfo("refreshToken from JWT = \n" + req.session.refreshTokenFromJWT);

            //let redirectUrl = req.session.decodedJWT.request.application.redirectUrl;
            //Utils.logInfo("Redirecting to: \n" + JSON.stringify(redirectUrl));
            //res.redirect(redirectUrl); // redirect to MC app landing page
        }
        catch(error)
        {
            let errorMsg = "Error while decoding JWT. Message: " + error;
            Utils.logError(errorMsg);
            res.status(400).send(errorMsg);
        }
    }
   
    /**
     * POST handler for: /logout
     * logout: called by Marketing Cloud when user logs out
     * 
     */
    public logout(req: express.Request, res: express.Response)
    {
        let sessionId = req.session.id;
        Utils.logInfo("logout called. SessionId = " + sessionId);

        // Clear out JWT and everything we got from it.
        req.session.jwtFromSFMC = "";
        req.session.decodedJWT = "";
        req.session.refreshTokenFromJWT = "";
        req.session.oauthAccessToken = "";
        req.session.oauthAccessTokenExpiry = "";
    
        //return res.redirect('https://mc.s11.exacttarget.com/'); // accepted
	    let self = this;
        Utils.logInfo("Logout");
       
    }

    /**
     * GET handler for: /appdemooauthtoken
     * getOAuthAccessToken: called by demo app to get an OAuth access token
     * 
     * More info: https://developer.salesforce.com/docs/atlas.en-us.noversion.mc-getting-started.meta/mc-getting-started/get-access-token.htm
     * 
     */
    public getOAuthAccessToken(req: express.Request, res: express.Response)
    {
        let self = this;
        let sessionId = req.session.id;
        let clientId = process.env.DF18DEMO_CLIENTID;
        let clientSecret = process.env.DF18DEMO_CLIENTSECRET;
        
        //console.log('here we are cid' + clientId);
        //console.log('here we are cs' + clientSecret);
        
        

        req.session.oauthAccessToken = "";
        req.session.oauthAccessTokenExpiry = "";

        Utils.logInfo("getOAuthAccessToken route entered. SessionId = " + sessionId);

        if (clientId && clientSecret)
        {
            Utils.logInfo("Getting OAuth Access Token with ClientID and ClientSecret from in environment variables.");
			Utils.logInfo("This was called from axios reactjs");
            
            // set the desired timeout in options
            

            
             self._apiHelper.getOAuthAccessToken(clientId, clientSecret)           
            .then((result) => {
                req.session.oauthAccessToken = result.oauthAccessToken;
                req.session.oauthAccessTokenExpiry = result.oauthAccessTokenExpiry;
                res.status(result.status).send(result.statusText);
                req.setTimeout(0, ()=>{});
		})
            .catch((err) => {
                res.status(500).send(err);
            });
        }
        else
        {
            // error
            let errorMsg = "ClientID or ClientSecret *not* found in environment variables."; 
            Utils.logError(errorMsg);
            res.status(500).send(errorMsg);
        }
		
		//Utils.logInfo("called for data extension creation");
		//		self.loadData(req, res);
	}
	
	public getCategoryIDforApp(req: express.Request, res: express.Response)
    {
        let self = this;
        Utils.logInfo("Get Category ID route method called.");
        self._apiHelper.getCategoryID(req, res);
		//Utils.logInfo("Request Body." + req);
    }
	
	public loadDataForApp(req: express.Request, res: express.Response)
    {
        let self = this;
        Utils.logInfo("loadData route entered.");
        self._apiHelper.loadData(req, res);
    }
	
	public loadDataForPage2App(req: express.Request, res: express.Response)
    {
        let self = this;
        Utils.logInfo("loadData route entered.");
        self._apiHelper.loadDataForPage2(req, res);
    }
	public createDataExtensionForApp(req: express.Request, res: express.Response)
    {
        let self = this;
        Utils.logInfo("Creating Data extension Dynamically.");
        self._apiHelper.createDataExtension(req, res);
		Utils.logInfo("Request Body." + req);
    }
	
	public DataExtensionFolderCheck(req: express.Request, res: express.Response){
		let self = this;
        Utils.logInfo("Checking the data extension folder and Data extension is available or not.");
        self._apiHelper.DataExtensionFolderCheck(req, res);
	}
	
	public orgSetupCheck(req: express.Request, res: express.Response){
		let self = this;
		Utils.logInfo("Org Setup check Routed to Helper class.");
        self._apiHelper.orgSetupCheck(req, res);
	}
	
	public dataExtensionTemplateCheck(req: express.Request, res: express.Response){
		let self = this;
        Utils.logInfo("Data Extension Template Routed to Helper class.");
        self._apiHelper.dataExtensionTemplateCheck(req, res);
	}
	
	public AppUserInfo(req: express.Request, res: express.Response){
		let self = this;
        Utils.logInfo("Checking the data extension folder and Data extension is available or not.");
        self._apiHelper.AppUserInfo(req, res);
	}
}
