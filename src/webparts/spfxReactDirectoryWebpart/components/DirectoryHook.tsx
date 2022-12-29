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
} from "office-ui-fabric-react";
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
    
       const _loadAlphabets = ():void => {
         const alphabets: string[] = [];
         for (let i = 65; i < 91; i++) {
           alphabets.push(String.fromCharCode(i));
         }
         setaz(alphabets);
       };
       

       console.log("alphabets",az)
       useEffect(() => {
         _loadAlphabets();
         setstate({...state})
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
             // onSearch={_searchUsers}
              value={state.searchText}
              //onChange={_searchBoxChanged}
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
              {//!pagedItems || pagedItems.length == 0 ? (
                }
            </>
          )}
        </>
      )}
    </div>
  );
};

export default DirectoryHook