import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import CrudTest from './components/CrudTest';
import { ICrudTestProps } from './interfaces/ICrudTestProps';

export interface ICrudTestWebPartProps {
}

export default class CrudTestWebPart extends BaseClientSideWebPart<ICrudTestWebPartProps> {


  public render(): void {
      const element: React.ReactElement<ICrudTestProps> = React.createElement(CrudTest,{spcontext: this.context});
      ReactDom.render(element, this.domElement);

    
  }
}