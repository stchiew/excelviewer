import * as React from 'react';
import styles from './ExcelViewer.module.scss';
import { IExcelViewerProps } from './IExcelViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ExcelViewer extends React.Component<IExcelViewerProps, {}> {

  private endpoint: string = "https://pandalenses.sharepoint.com/_vti_bin/ExcelRest.aspx/Shared%20Documents/chart.xlsx/model/charts('Chart%201')";
  public render(): React.ReactElement<IExcelViewerProps> {
    return (
      <div className={styles.excelViewer}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <img src={this.endpoint} />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
