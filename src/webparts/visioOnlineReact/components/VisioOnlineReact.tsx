import * as React from 'react';
import styles from './VisioOnlineReact.module.scss';
import { IVisioOnlineReactProps } from './IVisioOnlineReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class VisioOnlineReact extends React.Component<IVisioOnlineReactProps, {}> {
  public render(): React.ReactElement<IVisioOnlineReactProps> {
    return (
      <div className={styles.visioOnlineReact}>
        <div id='iframeHost' className={styles.iframeHost}></div>
      </div>
    );
  }

  public componentDidMount() {
    if (this.props.documentUrl) {
      this.props.visioService.load(this.props.documentUrl, this.props.zoomLevel);
    }
  }

  public async componentDidUpdate(prevProps: IVisioOnlineReactProps) {
    if ((this.props.documentUrl && (this.props.documentUrl !== prevProps.documentUrl)) ||
      (this.props.zoomLevel != prevProps.zoomLevel)) {
      this.props.visioService.load(this.props.documentUrl, this.props.zoomLevel);
    }
    if ((this.props.bHighLight !== prevProps.bHighLight) || (this.props.shapeName !== prevProps.shapeName)) {
      this.props.visioService.highlightShape(this.props.shapeName, this.props.bHighLight);
    }
    if ((this.props.bOverlay !== prevProps.bOverlay) || (this.props.shapeName !== prevProps.shapeName)) {
      this.props.visioService.addOverlay(this.props.shapeName, this.props.bOverlay);
    }    
  }

}
