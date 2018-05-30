import * as React from 'react';
import logo from './logo.svg';
import { UserAgentApplication,User } from 'msal';
import './App.css';
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
  private usersList:any[];
  private user:User;

  constructor(props: any) {
    super(props);
    this.clientapplication = new UserAgentApplication(applicationConfig.clientID, null, (errorDesc: string, token: string, error: string, tokenType: string) => {
      console.log(error);
    }, { cacheLocation: "localStorage" });

    // we cant use set state with out loading of the component, so state intilization. here i am checking if the user had previously logged in and token is stored in the local storage.
     this.user = this.clientapplication.getUser();
    if (this.user) {
      this.state = {
        loggedin: true,
        loggedinuser: this.user
      };
      this._getAllUsers();
    }
    else {
      this.state = {
        loggedin: false,
        loggedinuser: null
      };
    }
    this._loginpopup = this._loginpopup.bind(this);
    this._getAllUsers = this._getAllUsers.bind(this);
  }

  public render() {

    return (
      <div className="App">
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <h1 className="App-title">Welcome to React</h1>
        </header>
        {this.state.loggedin ? <span>Fetch the details <UsersList users={this.usersList} /> </span> : <input type="button" value="Login" onClick={this._loginpopup} />}
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
          this.usersList=res.value;
          
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

class UsersList extends React.Component<any,any>{

  render(){
    debugger;
    let users=this.props.users;
    return (
      <span>of the users {users}</span>
    );
  }
}

