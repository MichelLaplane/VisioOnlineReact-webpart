import * as React from 'react';
import styles from './VisioOnlineReact.module.scss';
import { IVisioOnlineReactProps } from './IVisioOnlineReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class VisioOnlineReact extends React.Component<IVisioOnlineReactProps, {}> {
  public render(): React.ReactElement<IVisioOnlineReactProps> {
    return (
            <div className={ styles.visioOnlineReact}>
              <div id='iframeHost' className={styles.iframeHost}></div>        
            </div>      
    );
  }

    public componentDidMount() {
        if (this.props.documentUrl) {
          this.props.visioService.load(this.props.documentUrl);
        }
      }
    
}
