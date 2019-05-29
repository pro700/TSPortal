import * as React from 'react';
import * as $ from 'jquery';
import styles from './Wp1.module.scss';
import { IWp1Props } from './IWp1Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Wp1 extends React.Component<IWp1Props, {}> {
  public render(): React.ReactElement<IWp1Props> {
    return (
      <div className={ styles.wp1 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
