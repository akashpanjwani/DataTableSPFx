import * as React from 'react';
import styles from './IframeDialog.module.scss';
import { IIframeDialogProps } from './IIframeDialogProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, DialogType } from 'office-ui-fabric-react';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/controls/IFrameDialog/IFrameDialog";
import * as $ from 'jquery';

export interface IIframeDialogState {
  isDlgOpen: boolean;
}
let isDlgOpen = true;
export class IframeDialog extends React.Component<IIframeDialogProps, IIframeDialogState> {
  public constructor(props: IIframeDialogProps) {
    super(props);

    this.state = {
      isDlgOpen: true
    };
  }

  public async componentDidMount(): Promise<void> {
    // Parse our comments, return if we have no comments
    isDlgOpen = this.props.isDlgOpen;

  }

  public async componentWillReceiveProps(nextProps: any) {
    isDlgOpen = nextProps.isDlgOpen;
    this.setState({
      isDlgOpen: nextProps.isDlgOpen
    });
  }

  public render(): React.ReactElement<IIframeDialogProps> {
    return (
      <div className={styles.iframeDialog}>

        <div className={styles.column}>
        {!isDlgOpen ?<IFrameDialog
            url={this.props.docEditUrl}
            hidden={false}
            width={'800px'}
            height={'600px'}
            modalProps={{
              isBlocking: true
            }}
            iframeOnLoad={iframe => {
              const windowClose = iframe.contentWindow.close;
              const windowClosingEvent = new Event('closeWindow');

              iframe.addEventListener('closeWindow', () => {
                this.setState({
                  isDlgOpen: false
                });
              });

              iframe.contentWindow.close = () => {
                iframe.dispatchEvent(windowClosingEvent);
                windowClose();
              };
            }}
            dialogContentProps={{
              type: DialogType.close,
              showCloseButton: true
            }}
            /*scrolling={'no'}
            seamless={false}
            allowFullScreen={true}
            name={'docFrame'}*/
            containerClassName={ 'ms-dialogMainOverride ' + styles.textDialog}
            onDismiss={() => { this._onDlgDismiss(); }}
          />:null}
          </div>

      </div>
    );
  }

  private _onClick(): void {
    this.setState({
      isDlgOpen: true
    });
  }

  private onCloseClick = () => {
    if (window.parent !== window) {
      window.close();
    }
  }

  private _onDlgDismiss(): void {
    this.setState({
      isDlgOpen: false
    });
    isDlgOpen = true;
    this.props.callback("reload");
  }

  private _onDlgLoaded(): void {
    console.log('dlg is loaeded');
  }
}