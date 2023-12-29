/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./PersonaCard.module.scss";
import { IPersonaCardProps } from "./IPersonaCardProps";
import { IPersonaCardState } from "./IPersonaCardState";
import { Log, Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";

import { Persona, PersonaSize, DocumentCard, DocumentCardType, Icon, FocusZone } from "office-ui-fabric-react";
import { SelectLanguage } from "../SelectLanguage";
import { callApiWithToken } from "../../../../fetch";
import { msalConfig, protectedResources } from "../../../../authConfig";
import { getClaimsFromStorage } from "../../../../utils/storageUtils";
import ChatService from "../SPServices/ChatService";

const EXP_SOURCE: string = "SPFxDirectory";
const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408"; //component Id of the Live Persona Card module

export class PersonaCard extends React.Component<IPersonaCardProps, IPersonaCardState> {
  constructor(props: IPersonaCardProps) {
    super(props);
    this.state = { pictureUrl: undefined };
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    const strings = SelectLanguage(this.props.prefLang);

    const CreateChat = (userId: string) => {
      console.log("CreateChat userId", userId);

      const activeAccount: any = this.props.activeAccount;
      const instance = this.props.instance;
      const resource = new URL(protectedResources.apiChat.endpoint).hostname;
      const request = {
        scopes: protectedResources.scopes.chatCreate,
        account: activeAccount,
        claims:
          activeAccount &&
          getClaimsFromStorage(`cc.${msalConfig.auth.clientId}.${activeAccount.idTokenClaims.oid}.${resource}`)
            ? window.atob(
                getClaimsFromStorage(`cc.${msalConfig.auth.clientId}.${activeAccount.idTokenClaims.oid}.${resource}`)
              )
            : undefined,
        sid: ChatService.context.pageContext.legacyPageContext.aadSessionId,
      };

      //let accessToken: string = "";
      instance
        .acquireTokenSilent(request)
        .then((response: { accessToken: any }) => {
          const accessToken = response.accessToken;
          callApiWithToken(accessToken, protectedResources.apiChat.endpoint, activeAccount, userId)
            .then((response) => {
              this.props.profileProperties.Chat = ChatService.fixUrl(response);
              this.componentDidMount();
            })

            .catch((error) => {
              console.log("callApiWithToken error", error);
            });
        })
        .catch((error: any) => {
          console.log("acquireTokenSilent error", error);
        });
    };

    return (
      <DocumentCard className={styles.documentCard} type={DocumentCardType.normal}>
        <div className={styles.persona}>
          <Persona
            text={this.props.profileProperties.DisplayName}
            secondaryText={this.props.profileProperties.Title}
            tertiaryText={this.props.profileProperties.Department}
            imageUrl={this.props.profileProperties.PictureUrl}
            size={PersonaSize.size48}
            imageShouldFadeIn={false}
            imageShouldStartVisible={true}
          >
            {this.props.profileProperties.Email ? (
              <div className={styles.textOverflow}>
                <Icon iconName="PublicEmail" style={{ fontSize: "12px", verticalAlign: "sub" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {
                    <a
                      aria-label={`mailto link ${this.props.profileProperties.Email} `}
                      href={`mailto: ${this.props.profileProperties.Email}`}
                    >
                      {this.props.profileProperties.Email}
                    </a>
                  }
                </span>
              </div>
            ) : (
              ""
            )}

            {this.props.profileProperties.Chat ? (
              <div className={styles.textOverflow}>
                <Icon iconName="Chat" style={{ fontSize: "12px", verticalAlign: "sub" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {
                    <a
                      aria-label={`link to chat`}
                      href={`${this.props.profileProperties.Chat}`}
                      target="_blank"
                      rel="noreferrer"
                    >
                      {strings.StartChatLabel}
                    </a>
                  }
                </span>
              </div>
            ) : (
              <div className={styles.textOverflow}>
                <Icon iconName="Chat" style={{ fontSize: "12px", verticalAlign: "sub" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {
                    <a
                      aria-label={`create chat`}
                      href="#"
                      onClick={() => {
                        CreateChat(this.props.profileProperties.Id);
                      }}
                    >
                      {strings.StartChatLabel}
                    </a>
                  }
                </span>
              </div>
            )}

            {this.props.profileProperties.WorkPhone ? (
              <div>
                <Icon iconName="Phone" style={{ fontSize: "12px" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}> {this.props.profileProperties.WorkPhone}</span>
              </div>
            ) : (
              ""
            )}
            {this.props.profileProperties.Location ? (
              <div className={styles.textOverflow}>
                <Icon iconName="Poi" style={{ fontSize: "12px" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}> {this.props.profileProperties.Location}</span>
              </div>
            ) : (
              ""
            )}
          </Persona>
        </div>
      </DocumentCard>
    );
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    return <div className={styles.personaContainer}>{<FocusZone>{this._PersonaCard()}</FocusZone>}</div>;
  }
}
