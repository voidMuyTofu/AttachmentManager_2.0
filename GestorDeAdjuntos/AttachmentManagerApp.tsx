import * as React from "react";
import { IFileItem, ItemList } from "./itemList";

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

export interface IAttachmentProps {
  regardingObjectId: string;
  regardingEntityName: string;
  files: IFileItem[];
  onAttach: (selectedFiles: IFileItem[]) => Promise<void>;
  context: ComponentFramework.Context<IInputs>;
}

export interface IAttachmentState {
  files: IFileItem[];
  columns: IColumn[];
  hiddenModal: boolean;
  isInProgress: boolean;
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
    this.selection = new Selection();
    this.allFiles.setItems(this.props.files);

    this.state = {
      files: this.allFiles.getItems(),
      hiddenModal: true,
      isInProgress: false,
      columns: this.allFiles.getColumns(),
    };

    this.attachFilesClicked = this.attachFilesClicked.bind(this);
    this.onFilterChanged = this.onFilterChanged.bind(this);
    this.onAttachClicked = this.onAttachClicked.bind(this);
    this.hideDialog = this.hideDialog.bind(this);
    this.onClickExaminar = this.onClickExaminar.bind(this);
    this.onChangeFileUnput = this.onChangeFileUnput.bind(this);
    this.getMultipleDocs = this.getMultipleDocs.bind(this);
  }

  public render(): React.JSX.Element {
    const { hiddenModal: hiddenDialog, files, columns } = this.state;
    return (
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
                    <DefaultButton
                      text="Adjuntar"
                      onClick={this.onAttachClicked}
                    />
                  </Stack.Item>
                  <Stack.Item grow align="stretch">
                    <SearchBox
                      styles={{ root: { width: "100%" } }}
                      placeholder="Buscar archivo"
                      onChange={this.onFilterChanged}
                    />
                  </Stack.Item>
                </Stack>
                <Stack>
                  {this.state.isInProgress && (
                    <ProgressIndicator
                      label="En progreso"
                      description="Copiando archivos de Sharepoint al correo"
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
              <Stack>
                <PrimaryButton
                  text="Examinar"
                  onClick={this.getMultipleDocs}
                ></PrimaryButton>
              </Stack>
            </ScrollablePane>
          </div>
        </Dialog>
      </div>
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

  private onClickExaminar(e?: any): void {}

  private onChangeFileUnput(e?: any): void {
    const fileUpload = e.target.files[0];
  }

  private onItemInvoked(item: IFileItem): void {
    console.log("Item invoked: " + item.fileName);
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

  private getMultipleDocs(): void {
    let docArray: string[] = [];
    try {
      let fileOptions = {
        accept: "image",
        allowMultipleFiles: true,
        maximumAllowedFileSize: 10000000,
      };
      this.props.context.device
        .pickFile(fileOptions)
        .then((file: ComponentFramework.FileObject[]) => console.log("asdsd"))
        .catch((error: any) => {
          console.log(error.message);
        });
    } catch (error: any) {
      console.log("getMultipleDocs " + error.message);
    }
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
