import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { callApiWithToken } from "../../../../fetch";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { InteractionType } from '@azure/msal-browser';

export class ChatServiceManager {
    public context: WebPartContext;

    public setup(context: WebPartContext): void {
      this.context = context;
    }
  
    public fixUrl(chatUrl: string): string {
      // force url to open in web browser
      return chatUrl.replace("/l/chat/", "/#/l/chat/");
    }

    public getChats(accessToken: string, activeAccount: any): Promise<MicrosoftGraph.Chat[]> {
        const apiTxt: string = 'https://graph.microsoft.com/v1.0/users/' + activeAccount.localAccountId + '/chats';
  
        return new Promise((resolve, reject) => {
          try {
            callApiWithToken(accessToken, apiTxt, activeAccount)
            .then((response) => {
                resolve(response.value);
            })
            .catch((error) => {
                console.log("error.message", error.message);
            });
          }
            catch(error) {
            console.log("error yo!", error);
            reject(error);
          }
        });
    }
}

const ChatService = new ChatServiceManager();
export default ChatService;