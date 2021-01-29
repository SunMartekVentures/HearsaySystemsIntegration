/*
 * Express App
 */

import * as express from "express";
import * as compression from "compression";  // compresses requests
import * as bodyParser from "body-parser";
import * as logger from "morgan";
import * as path from "path";
import * as favicon from "serve-favicon";
import * as session from "express-session";
import SfmcApiDemoRoutes from './SfmcApiDemoRoutes';
import SfmcAppDemoRoutes from './SfmcAppDemoRoutes';
import Utils from './Utils';



const PORT = process.env.PORT || 5000

// Create & configure Express server
const app = express();
//var MemoryStore = require('memorystore')(session)

// Express configuration
app.set("port", PORT);
app.set("views", path.join(__dirname, "../views"));
app.set('view engine', 'ejs');

// Use helmet. More info: https://expressjs.com/en/advanced/best-practice-security.html
var helmet = require('helmet')
app.use(helmet());
// Allow X-Frame from Marketing Cloud. Sets "X-Frame-Options: ALLOW-FROM http://exacttarget.com".
app.use(helmet.frameguard({
    action: 'allow-from',
    domain: 'https://mc.s11.exacttarget.com'
  }))

app.use(session({
    name: 'server-session-cookie-id',
    secret: 'sanagama-df18',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false }
}));

app.use(compression());
app.use(logger("dev"));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.raw({type: 'application/jwt'}));

// Setup static paths
app.use(express.static(path.join(__dirname, "../static")));
app.use(favicon(path.join(__dirname,'../static','images','favicons', 'favicon.ico')));

// Routes: pages


app.get('/', function(req, res) { Utils.initSampleDataAndRenderView(req, res, 'hearsayapi.ejs') });
app.get('/apidemo', function(req, res) { Utils.initSampleDataAndRenderView(req, res, 'hearsayapi.ejs') });
app.get('/appdemo', function(req, res) { Utils.initSampleDataAndRenderView(req, res, 'hearsayapp.ejs') });
app.get('/jbdemo', function(req, res) { Utils.initSampleDataAndRenderView(req, res, 'jbdemo.ejs') });

const apiDemoRoutes = new SfmcApiDemoRoutes();
const appDemoRoutes = new SfmcAppDemoRoutes();

// Routes: used by this demo app that internally call Marketing Cloud REST APIs
app.post('/apidemooauthtoken', function(req, res) {
  apiDemoRoutes.getOAuthAccessToken(req, res);
});

app.post('/loaddata', function(req, res) {
  apiDemoRoutes.loadData(req, res); });

app.post('/loaddatafortemp', function(req, res) {
  apiDemoRoutes.loadDataForPage2(req, res); });

app.post('/createde', function(req, res) {
  apiDemoRoutes.createDataExtension(req, res); });
    
// Routes: called when this demo app runs as a Marketing Cloud app in an IFRAME in the Marketing Cloud web UI
app.post('/appdemoauthtoken', function(req, res) {
  appDemoRoutes.getOAuthAccessToken(req, res); });

app.post('/loaddataforapp', function(req, res) {
  appDemoRoutes.loadDataForApp(req, res); });

app.post('/loaddatafortempapp', function(req, res) {
  appDemoRoutes.loadDataForPage2App(req, res); });

app.post('/createdeforapp', function(req, res) {
  appDemoRoutes.createDataExtensionForApp(req, res); });
  
  app.get('/getcategoryidforapp', function(req, res) {
  appDemoRoutes.getCategoryIDforApp(req, res); });
  
  //getOauthForCurrentBusinessUnit
  app.post('/getoauthforcurrentbusinessunit', function(req, res) {
  appDemoRoutes.getOauthForCurrentBusinessUnit(req, res); });
  
  //dataextensionfoldercheck
	app.post('/dataextensionfoldercheck', function(req, res) {
	appDemoRoutes.DataExtensionFolderCheck(req, res); });

//appuserinfo
	app.post('/appuserinfo', function(req, res) {
	appDemoRoutes.AppUserInfo(req, res); });

// Marketing Cloud POSTs the JWT to the '/login' endpoint when a user logs in
app.post('/login', function(req, res) {
  appDemoRoutes.login(req, res); });

// Marketing Cloud POSTs to the '/logout' endpoint when a user logs out
app.post('/logout', function(req, res) {
  appDemoRoutes.logout(req, res); });

module.exports = app;
