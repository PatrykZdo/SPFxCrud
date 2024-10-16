import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import CrudTest from './components/CrudTest';
import { ICrudTestProps } from './interfaces/ICrudTestProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IDropdownOption } from '@fluentui/react';

export interface ICrudTestWebPartProps {
  choices: string;
}

export default class CrudTestWebPart extends BaseClientSideWebPart<ICrudTestWebPartProps> {


  public render(): void {

    this.getCategoryChoiceFieldValue().then((choices: string[]) =>{
      const dropDownOptions: IDropdownOption[] = choices.map(choice => ({key: choice, text:choice}));
      const element: React.ReactElement<ICrudTestProps> = React.createElement(CrudTest,{spcontext: this.context, choices: dropDownOptions});
      ReactDom.render(element, this.domElement);
    }).catch((e)=>{console.log(e)});

    
  }

  private getCategoryChoiceFieldValue(): Promise<string[]>{
    const listName = "Products";
    const fieldName = "Category";

    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields/getbytitle('${fieldName}')`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<any> => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error('Błąd w odpowiedzi REST API');
        }
      })
      .then((fieldData: any): string[] => {
        if (fieldData && fieldData.Choices) {
          return fieldData.Choices;
        }
        throw new Error('Nie znaleziono wartości rozwijanej listy');
      })
      .catch((error: any) => {
        console.error('Błąd przy pobieraniu wartości z REST API:', error);
        console.log(JSON.stringify(error));
        return [];
      });
  }

}