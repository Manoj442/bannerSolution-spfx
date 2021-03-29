import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import './SpfxBannerCustomizer.scss';
import * as strings from 'SpfxBannerApplicationCustomizerStrings';
import "@pnp/polyfill-ie11";
import "core-js/stable/array/from";
import "core-js/stable/array/fill";
import "core-js/stable/array/iterator";
import "core-js/stable/promise";
import "core-js/stable/reflect";
import "es6-map/implement";
import "core-js/stable/symbol";
//import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from "@microsoft/sp-http";
const LOG_SOURCE: string = 'SpfxBannerApplicationCustomizer';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxBannerApplicationCustomizerProperties {
  // This is an example; replace with your own property 
  top:string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxBannerApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxBannerApplicationCustomizerProperties> {
    private _topPlaceholder: PlaceholderContent | undefined;
    private bannerData:{};
    
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);   
    sp.setup({   
      ie11:true,  
      sp:{
        baseUrl:"https://captureclicks.sharepoint.com/sites/CaptureClicks-Teams"
      }
    });
    let message: string = this.properties.top;
    if (!message) {
      message = '(No properties were provided.)';
    }       
    console.log(this.bannerData);   
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve<void>();
  }

private getBanner():Promise<any[]>{
  return sp.web.lists.getByTitle("BannerList").items.select("Title","EnableBanner","Status","BannerMessage2").get(); 
//   const urlValue:string= "https://captureclicks.sharepoint.com/sites/CaptureClicks-Teams/_api/Web/Lists/getbytitle('BannerList')/items?$Select=Title,EnableBanner,Status,BannerMessage2"; 
//   return this.context.spHttpClient.get(urlValue, SPHttpClient.configurations.v1)
// .then((data: SPHttpClientResponse) => data.json())
// .then((data: any) => { 
//   return data.value;
}

// private getBannerDetails:Promise<any[]>{
//   return sp.web.lists.getByTitle("BannerList").items.select("Title","EnableBanner","Status","BannerMessage2").get(); 
// }
  
private _renderPlaceHolders(): void {
  console.log("SpfxBannerApplicationCustomizer._renderPlaceHolders()");
  console.log(
    "Available placeholders: ",
    this.context.placeholderProvider.placeholderNames
      .map(name => PlaceholderName[name])
      .join(", ")
  );

  // Handling the top placeholder
  if (!this._topPlaceholder) {
    this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top,
      { onDispose: this._onDispose }
    );
  }

    // The extension should not assume that the expected placeholder is available.
    if (!this._topPlaceholder) {
      console.error("The expected placeholder (Top) was not found.");
      return;
    } 

  this.getBanner().then((data:any[])=>{
    this.bannerData = data.map(({Title,EnableBanner,BannerMessage2,Status})=>{
      const background = Status === 'Delta' ? '#A4262C' : Status === 'Live' ? '#007D34' : '#8F7034';
      const bannerTitle = Status === 'Delta' ? Title : BannerMessage2;
      const newUrl=`<a href="https://captureclicks.sharepoint.com/" target="_blank"> take me to the new site</a>`;
      this._topPlaceholder.domElement.innerHTML = EnableBanner ? `
        <div class="app">
          <div class="top" style="background-color:${background}">
            <i class="ms-icon ms-icon--info" aria-hidden="true"></i> ${bannerTitle} 
             ${Status === 'Live' ? newUrl: '' }
          </div>
        </div>`:null;
    }); 
  });
  
}

private _onDispose(): void {
  console.log('[SpfxBannerApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
}
}
