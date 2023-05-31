import { IColumn } from "@fluentui/react/lib/DetailsList";

export interface IFileItem {
  key: number | string;
  id: string;
  fileName: string;
  fileType: string;
  fileUrl: string;
  lastModifiedOn: Date;
  lastModifiedBy: string;
  iconclassname: string;
}

export class ItemList {
  private columns: IColumn[];
  private items: IFileItem[];

  constructor() {
    this.columns = [];
    this.items = [];

    this.setColumns();
  }

  private setColumns(): void {
    this.columns = [];
    this.columns.push({
      key: "iconclassname",
      name: "",
      fieldName: "iconclassname",
      minWidth: 20,
      maxWidth: 40,
      isResizable: false,
    });
    this.columns.push({
      key: "fileName",
      name: "Name",
      fieldName: "fileName",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    });
    this.columns.push({
      key: "fileType",
      name: "Type",
      fieldName: "fileType",
      minWidth: 50,
      maxWidth: 50,
      isResizable: true,
    });
    this.columns.push({
      key: "lastModifiedOn",
      name: "Last Modified On",
      fieldName: "lastModifiedOn",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    });
    this.columns.push({
      key: "lastModifiedBy",
      name: "Last Modified By",
      fieldName: "lastModifiedBy",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    });
  }

  public getColumns(): IColumn[] {
    return this.columns;
  }

  public getItems(): IFileItem[] {
    return this.items;
  }

  public setItems(items: IFileItem[]): void {
    if (items) this.items = items;
  }
}
