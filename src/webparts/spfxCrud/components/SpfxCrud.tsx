import * as React from "react";
import styles from "./SpfxCrud.module.scss";
import { ISpfxCrudProps } from "./ISpfxCrudProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

interface IListItem {
  ID: number;
  Title: string;
  Age: number;
}

interface IListItems {
  AllItems: IListItem[];

  //
  listTitle: string;
  listAge: number;
  listSelectedID: number;
}

export default class SpfxCrud extends React.Component<
  ISpfxCrudProps,
  IListItems
> {
  constructor(props: ISpfxCrudProps, state: IListItems) {
    super(props);
    this.state = {
      AllItems: [],
      listTitle: undefined,
      listAge: 0,
      listSelectedID: 0,
    };
  }
  componentDidMount() {
    this.getListItems();
  }
  // Get items
  public getListItems = () => {
    let listName = `Employee List`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    this.props.context.spHttpClient
      .get(requestURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      })
      .then((i) => {
        if (i == undefined) {
        } else {
          this.setState({
            AllItems: i.value,
          });
          console.log(this.state.AllItems);
        }
      });
  };

  // Delete item
  public deleteItem = (itemID: number) => {
    // alert("this is delete");
    let listName = `Employee List`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemID})`;

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=verbose",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item ID: ${itemID} deleted successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };

  // Add item
  public addItemInList = () => {
    // alert("this is delete");
    let listName = `Employee List`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const body: string = JSON.stringify({
      Title: this.state.listTitle,
      Age: this.state.listAge,
    });

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: body,
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item added successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };

  // Update item
  public updateItemInList = (itemID: number) => {
    // alert("this is delete");
    let listName = `Employee List`;

    let requestURL = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemID})`;

    const body: string = JSON.stringify({
      Title: this.state.listTitle,
      Age: this.state.listAge,
    });

    this.props.context.spHttpClient
      .post(requestURL, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: body,
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          alert(`Item updated successfully!`);
          this.getListItems();
        } else {
          alert(`Something went wrong!`);
          console.log(response.json());
        }
      });
  };
  public render(): React.ReactElement<ISpfxCrudProps> {
    return (
      <div className={styles["spfx-crud"]}>
        <input
          value={this.state.listTitle}
          type="text"
          name=""
          id="lsTitle"
          placeholder="Title"
          onChange={(e) => {
            this.setState({
              listTitle: e.currentTarget.value,
            });
            // console.log(this.state.listTitle);
          }}
        />
        <input
          value={this.state.listAge}
          type="number"
          name=""
          id="lsAge"
          placeholder="Age"
          onChange={(e) => {
            this.setState({
              listAge: e.currentTarget.value as any,
            });
          }}
        />
        <button
          onClick={() => {
            this.addItemInList();
          }}
        >
          Submit
        </button>
        <button
          onClick={() => {
            this.updateItemInList(this.state.listSelectedID);
          }}
        >
          Update
        </button>
        <hr />
        <hr />
        <table>
          <th>Title</th>
          <th>Age</th>
          {this.state.AllItems.map((emp) => {
            return (
              <tr>
                <td>{emp.Title}</td>
                <td>{emp.Age}</td>
                <td>
                  <button
                    onClick={() => {
                      this.setState({
                        listTitle: emp.Title,
                        listAge: emp.Age,
                        listSelectedID: emp.ID,
                      });
                    }}
                  >
                    Edit
                  </button>
                </td>
                <td>
                  <button
                    onClick={() => {
                      this.deleteItem(emp.ID);
                    }}
                  >
                    Delete
                  </button>
                </td>
              </tr>
            );
          })}
        </table>
      </div>
    );
  }
}
