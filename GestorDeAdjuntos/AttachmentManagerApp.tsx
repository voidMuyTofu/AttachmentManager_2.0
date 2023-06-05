import * as React from "react";
import { IFileItem, ItemList } from "./itemList";
import { createTheme, mergeStyleSets, ThemeProvider } from "@fluentui/react";

import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogType,
  Icon,
  MarqueeSelection,
  PrimaryButton,
  ProgressIndicator,
  ScrollablePane,
  ScrollbarVisibility,
  SearchBox,
  Stack,
  initializeIcons,
} from "@fluentui/react";

import {
  IColumn,
  Selection,
  IDetailsHeaderProps,
  DetailsList,
  DetailsListLayoutMode,
  ConstrainMode,
} from "@fluentui/react/lib/DetailsList";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { IRenderFunction } from "@fluentui/react/lib/Utilities";
import { classNames } from "./ComponentStyles";
import { IInputs } from "./generated/ManifestTypes";
import { EntityReference } from "./PCFHelper";
import { ActivityMimeAttachment } from "./Entity";

export interface IAttachmentProps {
  regardingObjectId: string;
  regardingEntityName: string;
  files: IFileItem[];
  onAttach: (selectedFiles: IFileItem[]) => Promise<void>;
  context: ComponentFramework.Context<IInputs>;
  primaryEntity: EntityReference;
  refreshAfterClose: () => void;
}

const myTheme = createTheme({
  palette: {
    themePrimary: "#e57e10",
    themeLighterAlt: "#fef9f5",
    themeLighter: "#fbe9d6",
    themeLight: "#f7d6b2",
    themeTertiary: "#f0af6a",
    themeSecondary: "#e98d2a",
    themeDarkAlt: "#cf720e",
    themeDark: "#ae600c",
    themeDarker: "#814709",
    neutralLighterAlt: "#faf9f8",
    neutralLighter: "#f3f2f1",
    neutralLight: "#edebe9",
    neutralQuaternaryAlt: "#e1dfdd",
    neutralQuaternary: "#d0d0d0",
    neutralTertiaryAlt: "#c8c6c4",
    neutralTertiary: "#a19f9d",
    neutralSecondary: "#605e5c",
    neutralSecondaryAlt: "#8a8886",
    neutralPrimaryAlt: "#3b3a39",
    neutralPrimary: "#323130",
    neutralDark: "#201f1e",
    black: "#000000",
    white: "#ffffff",
  },
});

export interface IAttachmentState {
  files: IFileItem[];
  columns: IColumn[];
  hiddenModal: boolean;
  isInProgress: boolean;
  attachmentButtonDisabled: boolean;
}

export class AttachmentManagerApp extends React.Component<
  IAttachmentProps,
  IAttachmentState
> {
  private selection: Selection;
  private allFiles: ItemList;
  private listadoArchivos: IFileItem[];

  constructor(props: IAttachmentProps) {
    super(props);

    initializeIcons();

    this.allFiles = new ItemList();
    this.selection = new Selection({
      onSelectionChanged: () => {
        debugger;
        this.handleSelection();
      },
    });

    this.allFiles.setItems(this.props.files);

    this.state = {
      files: this.allFiles.getItems(),
      hiddenModal: true,
      isInProgress: false,
      columns: this.allFiles.getColumns(),
      attachmentButtonDisabled: true,
    };

    this.attachFilesClicked = this.attachFilesClicked.bind(this);
    this.onFilterChanged = this.onFilterChanged.bind(this);
    this.onAttachClicked = this.onAttachClicked.bind(this);
    this.hideDialog = this.hideDialog.bind(this);
    this.onClickExaminar = this.onClickExaminar.bind(this);
    this.onChangeFileUnput = this.onChangeFileUnput.bind(this);
    this.handleSelection = this.handleSelection.bind(this);
    this.onItemInvoked = this.onItemInvoked.bind(this);
  }

  public render(): React.JSX.Element {
    const { hiddenModal: hiddenDialog, files, columns } = this.state;
    return (
      <ThemeProvider theme={myTheme}>
        <div>
          <CommandBar items={this.getItems()} />
          <Dialog
            hidden={hiddenDialog}
            onDismiss={() => {
              this.setState({ hiddenModal: true });
            }}
            dialogContentProps={{
              type: DialogType.close,
              title: "Adjuntar archivos",
              subText: "Selecciona los archivos que desee adjuntar al correo",
            }}
            modalProps={{ isBlocking: false }}
            minWidth="900px"
          >
            <div className={classNames.wrapper}>
              <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                <Sticky stickyPosition={StickyPositionType.Header}>
                  <Stack horizontal tokens={{ childrenGap: 20, padding: 10 }}>
                    <Stack.Item>
                      <PrimaryButton
                        className={classNames.buttonExaminar}
                        text="Adjuntar"
                        onClick={this.onAttachClicked}
                        disabled={this.state.attachmentButtonDisabled}
                      />
                    </Stack.Item>
                    <Stack.Item>
                      <DefaultButton
                        text="Adjuntar archivo local"
                        onClick={this.onClickExaminar}
                      ></DefaultButton>
                      <input
                        type="file"
                        id="fileInput"
                        multiple={true}
                        onChange={this.onChangeFileUnput}
                        style={{ display: "none" }}
                      />
                    </Stack.Item>
                    <Stack.Item grow align="stretch">
                      <SearchBox
                        styles={{ root: { width: "100%" } }}
                        placeholder="Buscar archivo"
                        onChange={this.onFilterChanged}
                        iconProps={{ style: { color: "#E57E10" } }}
                      />
                    </Stack.Item>
                  </Stack>
                  <Stack>
                    {this.state.isInProgress && (
                      <ProgressIndicator
                        label="En progreso"
                        description="Copiando archivos de Sharepoint al correo"
                        styles={{
                          progressBar: {
                            background:
                              "linear-gradient(90deg, rgba(237,235,233,1) 0%, rgba(229,126,16,1) 35%, rgba(237,235,233,1) 100%)",
                          },
                        }}
                      />
                    )}
                  </Stack>
                </Sticky>
                <MarqueeSelection selection={this.selection}>
                  <DetailsList
                    items={files}
                    columns={columns}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    constrainMode={ConstrainMode.unconstrained}
                    onRenderItemColumn={renderItemColumn}
                    onRenderDetailsHeader={onRenderDetailsHeader}
                    selection={this.selection}
                    selectionPreservedOnEmptyClick={true}
                    ariaLabelForSelectionColumn="Seleccionar archivo"
                    ariaLabelForSelectAllCheckbox="Seleccionar todos los archivos"
                    onItemInvoked={this.onItemInvoked}
                  />
                </MarqueeSelection>
              </ScrollablePane>
            </div>
          </Dialog>
        </div>
      </ThemeProvider>
    );
  }

  private getItems = () => {
    return [
      {
        key: "attachFile",
        name: "Click para adjuntar",
        cacheKey: "myCacheKey",
        iconProps: {
          iconName: "Attach",
        },
        ariaLabel: "Click para adjuntar",
        onClick: this.attachFilesClicked,
      },
    ];
  };

  private attachFilesClicked(): void {
    this.setState({ hiddenModal: false, isInProgress: false });
  }

  private onAttachClicked(): void {
    this.setState({ isInProgress: true });
    this.props.onAttach(this.getSelectedFiles()).then(this.hideDialog);
  }

  private hideDialog(): void {
    this.setState({ hiddenModal: true });
  }

  private onClickExaminar(e?: any): void {
    document.getElementById("fileInput")!.click();
  }

  private async onChangeFileUnput(e?: any) {
    const files = e.target.files;

    for (var i = 0; i < files.length; i++) {
      this.setState({ isInProgress: true });
      var content = await this.readAsync(files[i]);
      var fileName = files[i].name;
      var reader = new FileReader();
      var base64 = "";
      reader.readAsDataURL(files[i]);
      reader.onload = (readerEvent) => {
        var content = readerEvent.target!.result;
        base64 = (content as string).substring(
          (content as string).indexOf("base64,") + "base64,".length
        );
        ActivityMimeAttachment.create(
          base64,
          this.props.primaryEntity,
          fileName,
          this.props.context
        );
      };
    }
    this.props.refreshAfterClose();
    this.hideDialog();
  }

  private onItemInvoked(item: IFileItem): void {
    console.log("Item invoked: " + item.fileName);
  }

  private handleSelection(): void {
    const selectedFiles = this.getSelectedFiles();
    if (selectedFiles.length > 0) {
      console.log("Hay objetos");
      this.setState({ attachmentButtonDisabled: false });
    } else {
      console.log("No hay objetos");
      this.setState({ attachmentButtonDisabled: true });
    }
  }

  private async readAsync(file: any) {
    return new Promise((resolve: any, reject: any) => {
      const fileReader = new FileReader();
      fileReader.onload = (e) => {
        resolve(e.target!.result);
      };
      fileReader.readAsDataURL(file);
    });
  }

  private onFilterChanged(
    ev?: React.ChangeEvent<HTMLInputElement>,
    text?: string
  ): void {
    if (this.listadoArchivos == null) {
      this.listadoArchivos = this.state.files;
    }

    this.setState({
      files: text
        ? this.listadoArchivos.filter((item: IFileItem) => hasText(item, text))
        : this.listadoArchivos,
    });
  }

  private getSelectedFiles(): IFileItem[] {
    let selectedFiles: IFileItem[] = [];

    for (let i = 0; i < this.selection.getSelectedCount(); i++) {
      selectedFiles.push(this.selection.getSelection()[i] as IFileItem);
    }

    return selectedFiles;
  }
}

function hasText(item: IFileItem, text: string): boolean {
  return `${item.id}|${item.fileName}|${item.fileType}`.indexOf(text) > -1;
}

function onRenderDetailsHeader(
  props?: IDetailsHeaderProps,
  defaultRender?: IRenderFunction<IDetailsHeaderProps>
): React.JSX.Element {
  return (
    <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced={true}>
      {defaultRender && defaultRender({ ...props! })}
    </Sticky>
  );
}

function renderItemColumn(item: IFileItem, index?: number, column?: IColumn) {
  if (column) {
    const fieldContent = item[column.fieldName as keyof IFileItem] as string;

    switch (column.key) {
      case "iconclassname":
        return (
          <Icon iconName={fieldContent} className={classNames.fileIcon}></Icon>
        );
      case "lastModifiedOn":
        const dateField = item[column.fieldName as keyof IFileItem] as Date;
        return (
          <div>
            {dateField.toLocaleDateString("es-ES")}{" "}
            {dateField.toLocaleTimeString("es-ES")}
          </div>
        );
      case "fileType":
      case "fileName":
      case "lastModifiedBy":
        return <div>{fieldContent}</div>;
      default:
        break;
    }
  }
}
