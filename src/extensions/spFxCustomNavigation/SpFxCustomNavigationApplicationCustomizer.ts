import * as ReactDom from "react-dom";
import * as React from "react";

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { setup as pnpSetup } from '@pnp/common';

import * as strings from 'SpFxCustomNavigationApplicationCustomizerStrings';
import SideNav from "./components/SideNav/SideNav";

const LOG_SOURCE: string = 'SpFxCustomNavigationApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxCustomNavigationApplicationCustomizerProperties {
  // This is an example; replace with your own property
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxCustomNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFxCustomNavigationApplicationCustomizerProperties> {

  private sideNavPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    return super.onInit().then(() => {
      // render the extension only in teams
      if (navigator.userAgent.indexOf('Teams') === -1) {
        return;
      }

      const mainContentElement = document.getElementById('spPageChromeAppDiv');
      this.addClassToElement(mainContentElement, ['in-teams']);

      pnpSetup({ spfxContext: this.context });

      this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceholders);
    });
  }

  private renderPlaceholders(): void {
    if (!this.sideNavPlaceholder) {
      this.sideNavPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: this.onDispose
      });

      if (!this.sideNavPlaceholder) {
        return;
      }

      this.renderSideNav();
    }
  }

  private renderSideNav(): void {
    if (this.sideNavPlaceholder && this.sideNavPlaceholder.domElement) {
      const element: React.ReactElement<{}> = React.createElement(SideNav);
      ReactDom.render(element, this.sideNavPlaceholder.domElement);
    }
  }

  public onDispose(): void {
    console.log('[SideNav._onDispose] Disposed sidenav.');
    if (this.sideNavPlaceholder && this.sideNavPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(this.sideNavPlaceholder.domElement);
    }
    this.context.placeholderProvider.changedEvent.remove(this, this.renderPlaceholders);
  }

  private addClassToElement(element: HTMLElement, classNames: string[]) {
    classNames.forEach(className => {
      if (element && !element.classList.contains(className)) {
        element.classList.add(className);
      }
    });
  }
}
