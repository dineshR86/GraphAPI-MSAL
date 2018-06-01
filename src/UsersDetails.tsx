import * as React from 'react';
import { UserAgentApplication} from 'msal';
import * as GraphAPI from "@microsoft/microsoft-graph-client";
// import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
// import { List } from 'office-ui-fabric-react/lib/List';
import {IPersonaSharedProps,Persona} from 'office-ui-fabric-react/lib/Persona';

const examplePersona: IPersonaSharedProps = {
    optionalText: 'Available at 4:00pm',
    secondaryText: 'Designer',
    tertiaryText: 'In a meeting'
  };
  
//   const personaWithInitials: IPersonaSharedProps = {
//     ...examplePersona,
//     text: 'Maor Sharett',
//     imageInitials: 'MS'
//   };


export default class UsersDetails extends React.Component<any,any>{
    private clientapplication: UserAgentApplication;
    private graphClient:GraphAPI.Client;

    constructor(props:any){
        super(props);
        this.clientapplication=this.props.client;

        this.state={
            users:undefined
        }
    }

    public render(){
debugger;
        return (
            <div>
                Fetching the results
                <span>{this.clientapplication.getUser().name} </span>
                <div>
                 {/* <FocusZone direction={FocusZoneDirection.vertical}>
                    <List items={this.state.users} onRenderCell={this._onRenderCell} />
                 </FocusZone>    */}
                {this.state.users?this.state.users.map((value:any,index:number)=><Persona {...examplePersona} text={value.displayName} key={index} />): <span>Fetching</span>}
                 
                </div>
            </div>
        );
    }

    public componentDidMount(){
        if(!this.state.users){
         
        this.clientapplication.acquireTokenSilent(this.props.scopes,"",this.clientapplication.getUser()).then(
        (accesstoken)=>{
            this.graphClient= GraphAPI.Client.init({
                authProvider: (done) => {
                  done(null, accesstoken);
                }
            });

            // javascript wrapper for the graph API calls https://github.com/microsoftgraph/msgraph-sdk-javascript  
            this.graphClient.api('/users').get((err, res) => {
                debugger;
                if (err) {
                console.log("error from graph " + err); return;
                }
                else {
                console.log("Users " + res.value);
                this.setState({users:res.value});
                }
            });
        },
        (error)=>{
            debugger;
            console.log("error while fetching the access token silently"+error);
            this.clientapplication.acquireTokenPopup(this.props.scopes, "", this.clientapplication.getUser()).then((accesstoken) => {
                console.log("login popup" + accesstoken);
              }, (err) => {
                console.log("error" + err);
              });
        });
    }
    }

    // private _onRenderCell(item:any,index: number,isScrolling:boolean):any{
    //     debugger;
    //     return (
    //         <div>
    //             {item.displayName}
    //         </div>
    //     );
    // }
}