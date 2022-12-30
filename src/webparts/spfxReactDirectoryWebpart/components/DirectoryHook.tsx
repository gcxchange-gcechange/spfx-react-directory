//import  React from 'react'
import * as React from "react";
import { useState,useEffect } from "react";
import { IReactDirectoryProps } from './IReactDirectoryProps';
import { SelectLanguage } from "./SelectLanguage";
import { IReactDirectoryState } from "./IReactDirectoryState";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import styles from "./SpfxReactDirectoryWebpart.module.scss";
import {
  Pivot,
  PivotItem,
  PrimaryButton,
  SearchBox,
  Spinner,
  Stack,
  MessageBar,
  //Image,
  IStackTokens,
  IStackItemStyles,
  IStyleSet,
  IPivotStyles,
  PivotLinkFormat,
  PivotLinkSize,
  SpinnerSize,
  MessageBarType,
} from "@fluentui/react";

//import { SearchBox } from "@fluentui/react/lib/SearchBox";

//import ReactHtmlParser from "react-html-parser";

const DirectoryHook: React.FC<IReactDirectoryProps> = (props) => {
  const strings = SelectLanguage(props.prefLang);
  const [az, setaz] = useState<string[]>([]);
  // const [alphaKey, setalphaKey] = useState<string>("A");
  const [state, setstate] = useState<IReactDirectoryState>({
    users: [],
    isLoading: true,
    errorMessage: "",
    hasError: false,
    indexSelectedKey: "A",
    searchString: "FirstName",
    searchText: "",
  });
  // Paging
  //const [pagedItems, setPagedItems] = useState<any[]>([]);
  //const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  //const [currentPage, setCurrentPage] = useState<number>(1);

  const _loadAlphabets = (): void => {
    const alphabets: string[] = [];
    for (let i = 65; i < 91; i++) {
      alphabets.push(String.fromCharCode(i));
    }
    setaz(alphabets);
  };
  const _searchBoxChanged = (newvalue: string): void => {
    //setCurrentPage(1);
    setstate({
      ...state,
      searchText: newvalue,
    });
  };
let _searchUsers = async () => {
  try {
    setstate({
      ...state,
      isLoading: true,
    });
    const searchText = state.searchText;
    if (searchText.length > 0) {
      let searchProps: string[] =
        props.searchProps && props.searchProps.length > 0
          ? props.searchProps.split(",")
          : ["FirstName", "LastName", "PreferredName"];

      let qryText: string = "";
      let finalSearchText: string = searchText
        ? searchText.replace(/ /g, "+")
        : searchText;
      if (props.clearTextSearchProps) {
        let tmpCTProps: string[] =
          props.clearTextSearchProps.indexOf(",") >= 0
            ? props.clearTextSearchProps.split(",")
            : [props.clearTextSearchProps];
        if (tmpCTProps.length > 0) {
          searchProps.map((srchprop, index) => {
            let ctPresent: any[] = filter(tmpCTProps, (o) => {
              return o.toLowerCase() == srchprop.toLowerCase();
            });
            if (ctPresent.length > 0) {
              if (index == searchProps.length - 1) {
                qryText += `${srchprop}:${searchText}*`;
              } else qryText += `${srchprop}:${searchText}* OR `;
            } else {
              if (index == searchProps.length - 1) {
                qryText += `${srchprop}:${finalSearchText}*`;
              } else qryText += `${srchprop}:${finalSearchText}* OR `;
            }
          });
        } else {
          searchProps.map((srchprop, index) => {
            if (index == searchProps.length - 1)
              qryText += `${srchprop}:${finalSearchText}*`;
            else qryText += `${srchprop}:${finalSearchText}* OR `;
          });
        }
      } else {
        searchProps.map((srchprop, index) => {
          if (index == searchProps.length - 1)
            qryText += `${srchprop}:${finalSearchText}*`;
          else qryText += `${srchprop}:${finalSearchText}* OR `;
        });
      }
      console.log(qryText);
      const users = await _services.searchUsersNew(
        props.context,
        "",
        qryText,
        false
      );
      setstate({
        ...state,
        searchText: searchText,
        indexSelectedKey: null,
        users:
          users && users.PrimarySearchResults
            ? users.PrimarySearchResults
            : null,
        isLoading: false,
        errorMessage: "",
        hasError: false,
      });
    } else {
      setstate({ ...state, searchText: "" });
      //_searchByAlphabets(true);
    }
  } catch (err) {
    setstate({ ...state, errorMessage: err.message, hasError: true });
  }
};
  useEffect(() => {
    _loadAlphabets();
    setstate({ ...state });
    // _searchByAlphabets(true);
  }, [props]);

  const itemAlignmentsStackTokens: IStackTokens = {
    childrenGap: 20,
  };
  const stackItemStyles: IStackItemStyles = {
    root: {
      paddingTop: 5,
    },
  };
  const piviotStyles: Partial<IStyleSet<IPivotStyles>> = {
    link: {
      backgroundColor: "#ccc",
    },
  };

  return (
    <div className={styles.reactDirectory}>
      <WebPartTitle
        displayMode={props.displayMode}
        title={props.title}
        updateProperty={props.updateProperty}
      />
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
              //onSearch={_searchUsers}
              value={state.searchText}
              // onChange={() => _searchBoxChanged(state.searchText)}
              //onChange={_searchBoxChanged}
              onSearch={(newValu) => console.log("value is " + newValu)}
              onChange={(newValue) => _searchBoxChanged(newValue.target.value)}
            />
          </Stack.Item>
          <Stack.Item order={2}>
            <PrimaryButton
            //onClick={_searchUsers}
            >
              {strings.SearchButtonLabel}
            </PrimaryButton>
          </Stack.Item>
        </Stack>

        <div>
          {
            <Pivot
              styles={piviotStyles}
              className={styles.alphabets}
              linkFormat={PivotLinkFormat.tabs}
              selectedKey={state.indexSelectedKey}
              //onLinkClick={_alphabetChange}
              linkSize={PivotLinkSize.normal}>
              {az.map((index: string) => {
                return (
                  <PivotItem headerText={index} itemKey={index} key={index} />
                );
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
              <MessageBar messageBarType={MessageBarType.error}>
                {state.errorMessage}
              </MessageBar>
            </div>
          ) : (
            <>
              {
                //!pagedItems || pagedItems.length == 0 ? (
              }
            </>
          )}
        </>
      )}
    </div>
  );
};

export default DirectoryHook