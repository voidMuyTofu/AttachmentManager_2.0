import * as React from "react";
import ReactDOM = require("react-dom");
import { ActivityMimeAttachment, Email, SharePointDocument } from "./Entity";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { http } from "./http";
import { IconMapper } from "./iconMapper";
import { IFileItem } from "./itemList";
import {
  EntityReference,
  PrimaryEntity,
  isInHarness,
  SharePointHelper,
} from "./PCFHelper";
import { AttachmentManagerApp, IAttachmentProps } from "./AttachmentManagerApp";

export class GestorDeAdjuntos
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private container: HTMLDivElement;
  private context: ComponentFramework.Context<IInputs>;

  private primaryEntity: PrimaryEntity;
  private regardingId: string;
  private notifyOutputChanged: () => void;

  private iconMapper: IconMapper;
  private spHelper: SharePointHelper;

  constructor() {}

  private async onAttach(selectedFiles: IFileItem[]) {
    let apiUrl: string;
    for (let i = 0; i < selectedFiles.length; i++) {
      if (selectedFiles[i].fileUrl != null) {
        const fileUrl = selectedFiles[i].fileUrl;
        console.log(fileUrl);

        apiUrl = this.spHelper.makeApiUrl(fileUrl);

        const data = await http(apiUrl);
        ActivityMimeAttachment.create(
          data["Content"],
          this.primaryEntity.Entity,
          selectedFiles[i].fileName,
          this.context
        );
      }
    }
    this.refreshAfterClose();
  }

  private refreshAfterClose() {
    this.regardingId = new Date().toTimeString();
    this.notifyOutputChanged();
  }

  private renderControl(ec: ComponentFramework.WebApi.Entity[]): void {
    console.log("renderControl");
    let props: IAttachmentProps = {} as IAttachmentProps;
    props.files = [];
    props.onAttach = this.onAttach.bind(this);
    props.context = this.context;
    props.primaryEntity = this.primaryEntity.Entity;
    props.refreshAfterClose = this.refreshAfterClose.bind(this);

    for (let i = 0; i < ec.length; i++) {
      let file: IFileItem = {
        key: i,
        id: i.toString(),
        fileName: ec[i][SharePointDocument.FullName],
        fileUrl: ec[i][SharePointDocument.AbsoluteUrl],
        fileType: ec[i][SharePointDocument.FileType],
        iconclassname: this.iconMapper.getBySharePointIcon(
          ec[i][SharePointDocument.IconClassName]
        ),
        lastModifiedOn: new Date(ec[i][SharePointDocument.LastModifiedOn]),
        lastModifiedBy: ec[i][SharePointDocument.LastModifiedBy],
      };
      props.files.push(file);
    }

    ReactDOM.render(
      React.createElement(AttachmentManagerApp, props),
      this.container
    );
  }

  private renderControlWithMockData(): void {
    console.log("renderControl");
    let props: IAttachmentProps = {} as IAttachmentProps;
    ReactDOM.render(
      React.createElement(AttachmentManagerApp, props),
      this.container
    );
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;

    this.primaryEntity = new PrimaryEntity(this.context);
    this.iconMapper = new IconMapper();
    this.spHelper = new SharePointHelper(
      this.context.parameters.SharePointSiteURLs.raw as string,
      this.context.parameters.FlowURL.raw as string
    );
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;

    this.primaryEntity = new PrimaryEntity(this.context);

    if (isInHarness()) {
      this.renderControlWithMockData();
    } else {
      Email.getById(this.primaryEntity.Entity.id, this.context).then((e) => {
        const regarding: EntityReference = EntityReference.get(
          e,
          Email.RegardingObject
        );

        SharePointDocument.getByRegarding(
          regarding.id,
          regarding.typeName,
          this.context
        ).then((ec) => {
          console.log(`No. of documents in SP ${ec.length}`);
          this.renderControl(ec);
        });
      });
    }
  }

  public getOutputs(): IOutputs {
    return { RegardingId: this.regardingId };
  }

  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this.container);
  }
}
