import React from "react";
import { Button } from "reactstrap";
import withAuthProvider, { AuthComponentProps } from "../../common/AuthProvider";
import { config } from "../../Config";
import { createUsersChat, getCurrentUserId, getUserId, sendMessage } from './../../services/GraphService';

interface ChatState {
    messageSend: boolean;
  }

class ChatTest extends React.Component<AuthComponentProps, ChatState> {
    constructor(props: any) {
      super(props);
      this.state = {messageSend: false}
    }
  
    async onClick() {
        try {
            var accessToken = await this.props.getAccessToken(config.scopes);
  
            let currentUserId = await getCurrentUserId(accessToken);
            //let userIdToSendMessage = await getUserId(accessToken, "23f551e991fc92f9");
            let chatOfUser = await createUsersChat(accessToken, "23f551e991fc92f9", currentUserId);
            let result = await sendMessage(accessToken, chatOfUser, "First Try");
            console.log(result);
            this.setState({messageSend:true});
        }
        catch (err) {
            console.log(err);
        }
    }
    render() {
  
      return (
        <Button color="primary"
          className="mr-2"
          onClick={this.onClick.bind(this)}>Send</Button>
      );
    }
  }
  
  export default withAuthProvider(ChatTest);