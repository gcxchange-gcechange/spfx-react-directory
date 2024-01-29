/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { useEffect, useState } from "react";

import styles from "../ReactDirectory.module.scss";

import { IReactDirectoryState } from "../IReactDirectoryState";

import {
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
  IStackItemStyles,
  Image,
  IImageProps,
  ImageFit,
  PrimaryButton,
  Stack,
  IStackTokens,
} from "@fluentui/react";
import { IReactDirectoryProps } from "../IReactDirectoryProps";
import parse from "html-react-parser";
import { SelectLanguage } from "../SelectLanguage";
//import * as _ from "lodash";
import { ISPServices } from "../SPServices/ISPServices";
import { spservices } from "../SPServices/spservices";
import { PersonaCard } from "../PersonaCard/PersonaCard";
import Paging from "../Pagination/Paging";
import { useMsal, useMsalAuthentication } from "@azure/msal-react";
import { InteractionType } from "@azure/msal-browser";
import { getClaimsFromStorage } from "../../../../utils/storageUtils";
import { msalConfig, protectedResources } from "../../../../authConfig";
import ChatService from "../SPServices/ChatService";

const wrapStackTokens: IStackTokens = { childrenGap: 30 };

const PersonaCardMain: React.FC<IReactDirectoryProps> = (props) => {
  const { instance } = useMsal();
  const activeAccount: any = instance.getActiveAccount();
  const accountName: string = activeAccount
    ? activeAccount.username + " (" + activeAccount.localAccountId + ")"
    : "not active";

  const resource = new URL(protectedResources.apiChat.endpoint).hostname;
  const request = {
    scopes: protectedResources.scopes.chatRead,
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

  const { result } = useMsalAuthentication(InteractionType.Silent, { ...request, redirectUri: "" });

  const _getUserChats = async (accessToken: string, activeAccount: any) => {
    let chatUserId: string = "";
    // eslint-disable-next-line eqeqeq
    const connectedUserId: string = activeAccount != null ? activeAccount.localAccountId : null;
    let lookForUserId: string = "";
    let lookForUserName: string = "";
    let foundIt: boolean = false;
    const chatList: Chat[] = [];

    ChatService.getChats(accessToken, activeAccount).then((chatData) => {
      if (chatData) {
        console.log("chatData", chatData);

        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        const users = state.users;

        const n = users.PrimarySearchResults.length;
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        const o = chatData.length;

        for (let index = 0; index < n; index++) {
          lookForUserId = users.PrimarySearchResults[index].UniqueId;
          lookForUserName = users.PrimarySearchResults[index].Title;
          foundIt = false;
          const o = chatData.length;

          for (let idx = 0; idx < o; idx++) {
            chatUserId = chatData[idx].id.substring(3, 39);

            // eslint-disable-next-line eqeqeq
            if (chatUserId == connectedUserId) {
              chatUserId = chatData[idx].id.substring(40, 76);
            }

            // eslint-disable-next-line eqeqeq
            if (lookForUserId == chatUserId) {
              const chatUrl = ChatService.fixUrl(chatData[idx].webUrl);
              const chat: Chat = { userId: lookForUserId, displayName: lookForUserName, chatUrl: chatUrl };

              chatList.push(chat);
              foundIt = true;

              let user: any = users.PrimarySearchResults[index];
              user = {
                ...user,
                ChatURL: chatUrl,
              };
              users.PrimarySearchResults[index] = user;
            }
          }

          if (foundIt === false) {
            const chat: Chat = { userId: lookForUserId, displayName: lookForUserName, chatUrl: "" };
            chatList.push(chat);
          }
        }

        // eslint-disable-next-line @typescript-eslint/no-use-before-define
        setstate({
          // eslint-disable-next-line @typescript-eslint/no-use-before-define
          ...state,
          users: users,
        });
      }
    });

    console.log("chatList", chatList);
  };

  const strings: ISpfxReactDirectoryWebpartWebPartStrings = SelectLanguage(props.prefLang);
  let _services: ISPServices = null;
  _services = new spservices(props.context);

  const [az, setaz] = useState<string[]>([]);
  const [alphaKey, setalphaKey] = useState<string>("A");
  const [state, setstate] = useState<IReactDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: "",
    hasError: false,
    indexSelectedKey: "A",
    searchString: "FirstName",
    searchText: "",
    searchFinished: false,
  });
  const hidingUsers: string[] = props.hidingUsers && props.hidingUsers.length > 0 ? props.hidingUsers.split("/") : [];

  // Paging
  const [pagedItems, setPagedItems] = useState<unknown[]>([]);
  const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [startItem, setStartItem] = useState<number>(0);
  const [pgNo, setPgNo] = useState<number>(0);

  const _onPageUpdate = async (pageno?: number) => {
    if (pageno) {
      setPgNo(pageno);
    } else {
      setPgNo(0);
    }

    // pageno ? setPgNo(pageno) : setPgNo(0);
    const currentPge = pageno ? pageno : currentPage;
    setCurrentPage(currentPge);
    const startItemIndex = (currentPge - 1) * pageSize;
    setStartItem(startItemIndex);

    if (!pageno) {
      const filItems = state.users.PrimarySearchResults;
      setPagedItems(filItems);
      setstate({
        ...state,
        isLoading: false,
        searchFinished: true,
      });
    }
  };

  const _getCurrentPageUsers = async () => {
    if (pgNo > 0) {
      setstate({
        ...state,
        isLoading: true,
        searchFinished: false,
      });
      const searchText =
        state.searchText.length > 0 ? state.searchText : alphaKey.length > 0 && alphaKey !== "0" ? alphaKey : null;

      const users = await _services.searchUsersNew(
        props.context,
        `${searchText}`,
        "",
        true,
        hidingUsers,
        startItem,
        pageSize
        // AccessToken,
        // activeAccount
      );
      // setPagedItems(users.PrimarySearchResults);

      setstate({
        ...state,
        searchText: state.searchText,
        indexSelectedKey: state.indexSelectedKey,
        users: users && users.PrimarySearchResults ? users : null,
        // isLoading: false,
        errorMessage: "",
        hasError: false,
        // searchFinished: true,
      });
    }
  };

  const diretoryGrid =
    pagedItems && pagedItems.length > 0
      ? pagedItems.map((user: any, index: number) => {
          return (
            <PersonaCard
              key={index}
              context={props.context}
              prefLang={props.prefLang}
              activeAccount={activeAccount}
              instance={instance}
              // accessToken={""}
              // activeAccount={null}

              profileProperties={{
                Id: user.UniqueId,
                DisplayName:
                  user.FirstName && user.LastName ? `${user.FirstName}   ${user.LastName}` : user.PreferredName,
                Title: user.JobTitle,
                PictureUrl: user.PictureURL,
                Email: user.WorkEmail,
                Chat: user.ChatURL,
                Department: user.Department,
                WorkPhone: user.WorkPhone,
                Location: user.OfficeNumber ? user.OfficeNumber : user.BaseOfficeLocation,
              }}
            />
          );
        })
      : [];

  const _loadAlphabets = () => {
    const alphabets: string[] = [];
    for (let i = 65; i < 91; i++) {
      alphabets.push(String.fromCharCode(i));
    }
    setaz(alphabets);
  };

  const _alphabetChange = async (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) => {
    if (alphaKey !== item.props.itemKey) {
      setstate({
        ...state,
        searchText: "",
        indexSelectedKey: item.props.itemKey,
        isLoading: true,
        searchFinished: false,
      });
      setalphaKey(item.props.itemKey);
      setCurrentPage(1);
      setStartItem(0);
      setPgNo(0);
    }
  };

  const _searchByAlphabets = async (initialSearch: boolean) => {
    setstate({ ...state, isLoading: true, searchText: "" });
    let users = null;
    if (initialSearch) {
      users = await _services.searchUsersNew(props.context, "a", "", true, hidingUsers, startItem, pageSize);
    } else {
      users = await _services.searchUsersNew(props.context, `${alphaKey}`, "", true, hidingUsers, startItem, pageSize);
    }

    setstate({
      ...state,
      searchText: "",
      indexSelectedKey: initialSearch ? "A" : state.indexSelectedKey,
      users: users && users.PrimarySearchResults ? users : null,
      //isLoading: false,
      errorMessage: "",
      hasError: false,
      //searchFinished: true,
    });
  };

  const _searchUsers = async () => {
    try {
      setstate({
        ...state,
        isLoading: true,
        searchFinished: false,
      });
      setalphaKey("");
      const searchText = state.searchText;
      if (searchText.length > 0) {
        const searchProps: string[] = ["FirstName", "LastName"];

        let qryText: string = "";
        const finalSearchText: string = searchText ? searchText.replace(/ /g, "+") : searchText;

        searchProps.map((srchprop, index) => {
          if (index === searchProps.length - 1) qryText += `${srchprop}:${finalSearchText}*`;
          else qryText += `${srchprop}:${finalSearchText}* OR `;
        });

        const users = await _services.searchUsersNew(
          props.context,
          "",
          qryText,
          false,
          hidingUsers,
          startItem,
          pageSize
          //   AccessToken,
          //   activeAccount
        );

        setstate({
          ...state,
          searchText: searchText,
          indexSelectedKey: null,
          users: users && users.PrimarySearchResults ? users : null,
          // isLoading: false,
          errorMessage: "",
          hasError: false,
          // searchFinished: true,
        });
      } else {
        setstate({ ...state, searchText: "" });
        await _searchByAlphabets(true);
      }
    } catch (err) {
      setstate({ ...state, errorMessage: err.message, hasError: true });
    }
  };

  const _searchBoxChanged = (newvalue: string): void => {
    setCurrentPage(1);
    setStartItem(0);
    setPgNo(0);
    setstate({
      ...state,
      searchText: newvalue,
      searchFinished: false,
    });
  };

  //_searchUsers = _.debounce(_searchUsers, 500);
  useEffect(() => {
    _loadAlphabets();
  }, []);
  useEffect(() => {
    if (alphaKey.length > 0 && alphaKey !== "0") {
      _searchByAlphabets(false)
        .then((data) => {
          return data;
        })
        .catch((err) => {
          /* perform error handling if desired */
        });
    }
  }, [alphaKey]);
  useEffect(() => {
    setPageSize(props.pageSize);
    if (state.users.PrimarySearchResults) {
      _onPageUpdate()
        .then((data) => {
          return data;
        })
        .catch((err) => {
          /* perform error handling if desired */
        });
    }
  }, [state.users]);

  useEffect(() => {
    if (pgNo > 0) {
      _getCurrentPageUsers()
        .then((data) => {
          return data;
        })
        .catch((err) => {
          /* perform error handling if desired */
        });
    }
  }, [pgNo]);

  useEffect(() => {
    //_loadAlphabets();

    _searchByAlphabets(true)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        /* perform error handling if desired */
      });
  }, [props]);

  useEffect(() => {
    if (state.searchFinished) {
      if (result) {
        _getUserChats(result.accessToken, activeAccount);
      }
    }
  }, [accountName, state.searchFinished]); // state.searchFinished

  const itemAlignmentsStackTokens: IStackTokens = {
    childrenGap: 20,
  };
  const stackItemStyles: IStackItemStyles = {
    root: {
      paddingTop: 5,
    },
  };

  const imageProps: Partial<IImageProps> = {
    imageFit: ImageFit.centerContain,
    width: 200,
    height: 200,
    src: require("../../assets/HidingYeti.png"),
  };

  // const piviotStyles: Partial<IStyleSet<IPivotStyles>> = {
  //   link: {
  //     backgroundColor: "#e3e1e1",
  //     color: "#000",
  //     fontSize: "17px",
  //   },
  //   linkIsSelected: {
  //     fontSize: "17px",
  //   },
  // };

  return (
    <div className={styles.reactDirectory} lang={props.prefLang} style={{minHeight:"300px"}}>
      <div className={styles.searchBox}>
        <Stack horizontal tokens={itemAlignmentsStackTokens}>
          <Stack.Item order={1} styles={stackItemStyles}>
            <span>
              <label>{strings.SearchBoxLabel}</label>
            </span>
          </Stack.Item>
          <Stack.Item order={2}>
            <SearchBox
              placeholder={strings.SearchPlaceHolder}
              className={styles.searchTextBox}
              onSearch={_searchUsers}
              value={state.searchText}
              onChanged={_searchBoxChanged}
            />
          </Stack.Item>
          <Stack.Item order={2}>
            <PrimaryButton onClick={_searchUsers}>{strings.SearchButtonLabel}</PrimaryButton>
          </Stack.Item>
        </Stack>

        <div>
          {
            <Pivot
              // styles={piviotStyles}
              className={styles.alphabets}
              linkFormat={PivotLinkFormat.tabs}
              selectedKey={state.indexSelectedKey}
              onLinkClick={_alphabetChange}
              linkSize={PivotLinkSize.normal}
            >
              {az.map((index: string) => {
                return <PivotItem headerText={index} itemKey={index} key={index} />;
              })}
            </Pivot>
          }
        </div>
      </div>
      {state.isLoading ? (
        <div style={{ marginTop: "10px" }}>
          <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
        </div>
      ) : (
        <>
          {state.hasError ? (
            <div style={{ marginTop: "10px" }}>
              <MessageBar messageBarType={MessageBarType.error}>{state.errorMessage}</MessageBar>
            </div>
          ) : (
            <>
              {!pagedItems || pagedItems.length === 0 ? (
                <div className={styles.noUsers}>
                  <Stack horizontal tokens={itemAlignmentsStackTokens}>
                    <Stack.Item order={1} styles={stackItemStyles}>
                      <span tabIndex={0}>
                        <Image {...imageProps} alt={strings.NoUserFoundImageAltText} />
                      </span>
                    </Stack.Item>
                    <Stack.Item order={2}>
                      <span>
                        <p tabIndex={0}>{parse(strings.DirectoryMessage)}</p>
                        <PrimaryButton href={strings.NoUserFoundEmail}>{strings.NoUserFoundLabelText}</PrimaryButton>
                      </span>
                    </Stack.Item>
                  </Stack>
                </div>
              ) : (
                <>
                  <div style={{ width: "100%", display: "inline-block" }}>
                    <Paging
                      totalItems={state.users.TotalRows}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage}
                    />
                  </div>

                  <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                    <div>{diretoryGrid}</div>
                  </Stack>
                  <div style={{ width: "100%", display: "inline-block" }}>
                    {
                      <Paging
                        totalItems={state.users.TotalRows}
                        itemsCountPerPage={pageSize}
                        onPageUpdate={_onPageUpdate}
                        currentPage={currentPage}
                      />
                    }
                  </div>
                </>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
};

export default PersonaCardMain;
