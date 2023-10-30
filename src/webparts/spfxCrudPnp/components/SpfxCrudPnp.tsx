import * as React from "react";
import styles from "./SpfxCrudPnp.module.scss";
import type { ISpfxCrudPnpProps } from "./ISpfxCrudPnpProps";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export default class SpfxCrudPnp extends React.Component<
  ISpfxCrudPnpProps,
  {}
> {
  public render(): React.ReactElement<ISpfxCrudPnpProps> {
    return (
      <div className={styles.spfxCrudPnp}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item ID:</div>
                <input type="text" id="itemId"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input type="text" id="fullName"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Age</div>
                <input type="text" id="age"></input>
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems"></div>
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>
                    Create
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>
                    Read
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>
                    Read All
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>
                    Update
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>
                    Delete
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  //Create Item
  private createItem = async () => {
    try {
      const fullNameInput = document.getElementById(
        "fullName"
      ) as HTMLInputElement;
      const ageInput = document.getElementById("age") as HTMLInputElement;

      const addItem = await sp.web.lists
        .getByTitle("EmployeeDetail")
        .items.add({
          Title: fullNameInput.value,
          Age: ageInput.value,
        });
      console.log(addItem);
      alert(`Item created successfully with ID: ${addItem.data.ID}`);
    } catch (e) {
      console.error(e);
    }
  };

  //Get Item by ID
  private getItemById = async () => {
    try {
      const idElement = document.getElementById('itemId') as HTMLInputElement;
      const id: number = Number(idElement.value);
  
      if (id > 0) {
        const item: any = await sp.web.lists.getByTitle("EmployeeDetail").items.getById(id).get();
        const fullNameElement = document.getElementById('fullName') as HTMLInputElement;
        const ageElement = document.getElementById('age') as HTMLInputElement;
  
        if (fullNameElement && ageElement) {
          fullNameElement.value = item.Title;
          ageElement.value = item.Age;
        } else {
          console.error("One or both elements not found.");
        }
      } else {
        alert(`Please enter a valid item id.`);
      }
    } catch (e) {
      console.error(e);
    }
  }
  
 
  
//Get all items
private getAllItems = async () => {
  try {
    const items: any[] = await sp.web.lists.getByTitle("EmployeeDetail").items.get();
    console.log(items);
    if (items.length > 0) {
      const allItemsElement = document.getElementById("allItems");
      if (allItemsElement) {
        let html = `<table><tr><th>ID</th><th>Full Name</th><th>Age</th></tr>`;
        items.forEach(item => {
          html += `<tr><td>${item.ID}</td><td>${item.Title}</td><td>${item.Age}</td></tr>`;
        });
        html += `</table>`;
        allItemsElement.innerHTML = html;
      } else {
        console.error("Element with id 'allItems' not found.");
      }
    } else {
      alert(`List is empty.`);
    }
  } catch (e) {
    console.error(e);
  }
}

 
  
// Update Item
private updateItem = async () => {
  try {
    const idElement = document.getElementById('itemId') as HTMLInputElement;
    const id: number = Number(idElement.value);

    if (id > 0) {
      const fullNameElement = document.getElementById("fullName") as HTMLInputElement;
      const ageElement = document.getElementById("age") as HTMLInputElement;

      if (fullNameElement && ageElement) {
        const itemUpdate = await sp.web.lists.getByTitle("EmployeeDetail").items.getById(id).update({
          'Title': fullNameElement.value,
          'Age': ageElement.value
        });
        console.log(itemUpdate);
        alert(`Item with ID: ${id} updated successfully!`);
      } else {
        console.error("One or both elements not found.");
      }
    } else {
      alert(`Please enter a valid item id.`);
    }
  } catch (e) {
    console.error(e);
  }
}

// Delete Item
private deleteItem = async () => {
  try {
    const idElement = document.getElementById('itemId') as HTMLInputElement;
    const id: number = Number(idElement.value);

    if (id > 0) {
      let deleteItem = await sp.web.lists.getByTitle("EmployeeDetail").items.getById(id).delete();
      console.log(deleteItem);
      alert(`Item ID: ${id} deleted successfully!`);
    } else {
      alert(`Please enter a valid item id.`);
    }
  } catch (e) {
    console.error(e);
  }
}

}
