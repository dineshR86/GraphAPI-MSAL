import * as React from 'react';
import logo from './logo.svg';
import { UserAgentApplication,User } from 'msal';
import './App.css';
import UsersDetails from './UsersDetails';
// import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import * as GraphAPI from "@microsoft/microsoft-graph-client";

const applicationConfig = {
  clientID: '84d5a8cc-6dd8-49f8-ac9f-6c32b0035092',
  graphEndpoint: "https://graph.microsoft.com/v1.0/me/sendMail",
  graphScopes: ["user.read", "user.readbasic.all"]
};


export default class App extends React.Component<any, any> {
  private clientapplication: UserAgentApplication;
  private graphClient: GraphAPI.Client;
  private user:User;

  constructor(props: any) {
    super(props);
    this.clientapplication = new UserAgentApplication(applicationConfig.clientID, null, (errorDesc: string, token: string, error: string, tokenType: string) => {
      console.log(error);
    }, { cacheLocation: "localStorage" });

  this.state = {
        loggedin: false,
  };
    
    this._loginpopup = this._loginpopup.bind(this);
    this._getAllUsers = this._getAllUsers.bind(this);
  }

  // we cant use set state with out loading of the component, so updating the state of the component in the componentdidmount event. 
  // We can also do it in the constructor of the component.
  public componentDidMount(){
    if(this.clientapplication.getUser()){
      this.setState({
        loggedin:true
      });
    }
  }

  public render() {

    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title">Welcome to React</h1>
        </header>
        {this.state.loggedin ? <UsersDetails client={this.clientapplication} scopes={applicationConfig.graphScopes} /> : <input type="button" value="Login" onClick={this._loginpopup} />}
      </div>
    );
  }

  private _loginpopup() {
    debugger;
    this.clientapplication.loginPopup(applicationConfig.graphScopes).then((idtoken) => {
      this.user=this.clientapplication.getUser();
      this.setState({ loggedin: true });
    }, (error) => {
      console.log("error on login popup" + error);
    });
  }

  private _getAllUsers() {
    this.clientapplication.acquireTokenSilent(applicationConfig.graphScopes, "", this.user).then((accesstoken) => {
      console.log("token silently" + accesstoken);
      this.graphClient = GraphAPI.Client.init({
        authProvider: (done) => {
          done(null, accesstoken);
        }
      });

      // javascript wrapper for the graph API calls https://github.com/microsoftgraph/msgraph-sdk-javascript  
      this.graphClient.api('/users').get((err, res) => {
        if (err) {
          console.log("error from graph " + err); return;
        }
        else {
          console.log("Users " + res.value);
        }
      });
    }, (error) => {
      this.clientapplication.acquireTokenPopup(applicationConfig.graphScopes, "", this.user).then((accesstoken) => {
        console.log("login popup" + accesstoken);
      }, (err) => {
        console.log("error" + err);
      });
    });
  }
}

