import * as React from "react";
import styles from "./SetPageLayout.module.scss";
import { ISetPageLayoutProps } from "./ISetPageLayoutProps";
import { escape } from "@microsoft/sp-lodash-subset";
import {
  autobind,
  TextField,
  ChoiceGroup,
  PrimaryButton,
  SearchBox,
  Shimmer,
  ShimmerElementsGroup,
  ShimmerElementType,
  Label,
  Calendar,
  DatePicker,
  MessageBar,
  MessageBarType
} from "office-ui-fabric-react";

export interface IPageState {
  webUrl: string;
  pageUrl: string;
  pageLayout: string;
  error: boolean;
  errorMessage: string;
}

export default class SetPageLayout extends React.Component<
  ISetPageLayoutProps,
  IPageState
> {
  public constructor(props: ISetPageLayoutProps, state: IPageState) {
    super(props);
    this.state = {
      webUrl: "",
      pageUrl: "",
      pageLayout: "Article",
      error: false,
      errorMessage: ""
    };
  }

  public render(): React.ReactElement<ISetPageLayoutProps> {
    const { webUrl, pageUrl, pageLayout, error, errorMessage } = this.state;

    return (
      <div className={styles.setPageLayout}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Set Page Layout for a page</span>
              {!error && errorMessage != "" ? (
                <MessageBar
                  messageBarType={MessageBarType.success}
                  isMultiline={false}
                >
                  {errorMessage}
                </MessageBar>
              ) : null}
              {error && errorMessage != "" ? (
                <MessageBar
                  messageBarType={MessageBarType.error}
                  isMultiline={false}
                >
                  {errorMessage}
                </MessageBar>
              ) : null}
              <br></br>
              <TextField
                label="Web Url: "
                placeholder="e.g. https://wawadev.sharepoint.com/sites/DEVSBX/Sub-Site/"
                onChange={this._changeWebUrl}
              ></TextField>
              <TextField
                label="Page Url: "
                placeholder="e.g. /sites/DEVSBX/Sub-Site/SitePages/page.aspx"
                onChange={this._changePageUrl}
              ></TextField>
              <ChoiceGroup
                label="Select Page Layout:"
                onChange={this._changePageLayout}
                defaultSelectedKey="Article"
                options={[
                  {
                    key: "Article",
                    text: "Article"
                  },
                  {
                    key: "SingleWebPartAppPage",
                    text: "SingleWebPartAppPage"
                  }
                ]}
              ></ChoiceGroup>
              <br></br>
              <PrimaryButton
                text="  Save Layout  "
                onClick={() =>
                  this._savePageLayout(this.state)
                    .then(isSuccess => {
                      this.setState({
                        error: !isSuccess,
                        errorMessage: isSuccess
                          ? "Page layout has been updated"
                          : "Error occured while updating page layout"
                      });
                    })
                    .catch(isError => {
                      this.setState({
                        error: isError,
                        errorMessage: "Error occured while updating page layout"
                      });
                    })
                }
              ></PrimaryButton>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _updateError(msg: string) {
    this.setState({
      errorMessage: msg
    });
  }

  @autobind
  private _changeWebUrl(e: any, newValue: string) {
    this.setState({
      webUrl: newValue,
      error: false,
      errorMessage: ""
    });
  }

  @autobind
  private _changePageUrl(e: any, newValue: string) {
    this.setState({
      pageUrl: newValue,
      error: false,
      errorMessage: ""
    });
  }

  @autobind
  private _changePageLayout(e: any, newValue: any) {
    this.setState({
      pageLayout: newValue.key,
      error: false,
      errorMessage: ""
    });
  }

  _savePageLayout = (state: IPageState): Promise<boolean> => {
    const { webUrl, pageUrl, pageLayout } = state;
    return new Promise<boolean>(
      (
        resolve: (json: boolean) => void,
        reject: (error: boolean) => void
      ): void => {
        this._checkFileExist(webUrl, pageUrl).then(isFileExist => {
          if (!isFileExist) reject(false);

          fetch(webUrl + "_api/contextinfo", {
            method: "POST",
            headers: {
              accept: "application/json;odata=nometadata"
            }
          })
            .then(function(response) {
              return response.json();
            })
            .then(function(ctx) {
              return fetch(
                webUrl +
                  "_api/web/getfilebyurl('" +
                  pageUrl +
                  "')/ListItemAllFields",
                {
                  method: "POST",
                  headers: {
                    accept: "application/json;odata=nometadata",
                    "X-HTTP-Method": "MERGE",
                    "IF-MATCH": "*",
                    "X-RequestDigest": ctx.FormDigestValue,
                    "content-type": "application/json;odata=nometadata"
                  },
                  body: JSON.stringify({
                    PageLayoutType: pageLayout
                  })
                }
              );
            })
            .then(function(res) {
              res.ok ? resolve(true) : reject(false);
            })
            .catch(error => {
              reject(false);
            });
        });
      }
    );
  };

  _checkFileExist = (webUrl: string, pageUrl: string): Promise<boolean> => {
    return new Promise<boolean>(
      (
        resolve: (json: boolean) => void,
        reject: (error: boolean) => void
      ): void => {
        fetch(webUrl + "_api/web/getfilebyurl('" + pageUrl + "')", {
          method: "GET"
        })
          .then(res => {
            if (!res.ok) {
              this.setState({
                error: true,
                errorMessage: "Error occured 'File not found.'"
              });
              throw new Error("File not found");
              reject(false);
            } else resolve(true);
          })
          .catch(() => {
            reject(false);
          });
      }
    );
  };
}
