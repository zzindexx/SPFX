import * as React from 'react';
import styles from './SiteQuota.module.scss';
import { ISiteQuotaProps, ISiteQuotaState } from './ISiteQuotaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { PopupWindowPosition } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneLink/IPropertyPaneLink';
import pnp from "sp-pnp-js";

export default class SiteQuota extends React.Component<ISiteQuotaProps, ISiteQuotaState> {
  constructor(props:ISiteQuotaProps) {
    super(props);
    this.state = {
      isLoading: false,
      percentageUsed: 0,
      totalSpace: 0,
      usedSpace: 0
    };
  }

  public async componentDidMount() {
    this.setState((current: ISiteQuotaState) => ({
      isLoading: true,
      percentageUsed: 0,
      totalSpace: 0,
      usedSpace: 0
    }));

    const usage:[number, number] = await this.getData();
    console.log(usage);
    this.setState((current: ISiteQuotaState) => ({
      isLoading: false,
      percentageUsed: usage[1],
      totalSpace: usage[0] / usage[1],
      usedSpace: usage[0]
    }));
  }

  private async getData(): Promise<[number, number]> {
    return new Promise<[number, number]>(
      (resolve, reject) => {
        pnp.sp.site.select('Usage').get().then((data) => {
          resolve([parseFloat(data.Usage.Storage), data.Usage.StoragePercentageUsed]);
        });
      }
    );
  } 

  private formatSize(sizeInBytes: number): string {
    if (sizeInBytes < 1024)
      return sizeInBytes + " Bytes";
    if (sizeInBytes < (1024 * 1024))
      return (sizeInBytes / 1024).toFixed(2) + " KBs";
    if (sizeInBytes < (1024 * 1024 * 1024))
      return (sizeInBytes / 1024 / 1024).toFixed(2) + " MBs";
    if (sizeInBytes < (1024 * 1024 * 1024 * 1024))
      return (sizeInBytes / 1024 / 1024 / 1024).toFixed(2) + " GBs";
  }

  public render(): React.ReactElement<ISiteQuotaProps> {
    let desc: string = `Used space: ${this.formatSize(this.state.usedSpace)} Total space: ${this.formatSize(this.state.totalSpace)}`;
    const indicator: JSX.Element = this.state.isLoading ? <ProgressIndicator title="Site quota usage" className={styles.progressIndicator} description={desc} /> : <ProgressIndicator percentComplete={this.state.percentageUsed} title="Site quota usage" className={styles.progressIndicator} description={desc} />;

    return (
      <div className={ styles.siteQuota }>
        {indicator}
      </div>
    );
  }
}
