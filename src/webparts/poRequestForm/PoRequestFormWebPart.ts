import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import { Version } from '@microsoft/sp-core-library';
import PORequestForm from './components/PoRequestForm';
import { IPoRequestFormProps } from './components/IPoRequestFormProps';
import { initializeIcons } from '@fluentui/react/lib/Icons';

/** PnP Property Controls: List Picker */
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

/** PnPjs (SP v3) */
import { SPFI, spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import '@pnp/sp/webs';
import '@pnp/sp/lists';

export interface IPORequestFormWebPartProps {
  /** Lista PADRE (PO Requests) */
  parentListId: string | null;
  parentListTitle?: string;

  /** Lista HIJA (PO Request Lines) */
  childListId: string | null;
  childListTitle?: string;

  /** Lista de Proveedores sugeridos */
  suppliersListId: string | null;
  suppliersListTitle?: string;

  /** ✅ Nueva lista hija de Suppliers (campos de texto, directa al PR) */
  supplierChildListId: string | null;
  supplierChildListTitle?: string;

  /** ✅ NUEVAS LISTAS PARA DROPDOWNS (Coding) */
  projectsListId: string | null;
  projectsListTitle?: string;

  glCodesListId: string | null;
  glCodesListTitle?: string;

  /** ✅ NUEVA LISTA PARA COMPANY (dropdown) */
  companiesListId: string | null;
  companiesListTitle?: string;

  authorizersListId: string | null;
  authorizersListTitle?: string;
  /** ✅ NUEVA LISTA PARA APPROVERS (reglas / aprobadores) */
  approversListId: string | null;
  approversListTitle?: string;

  /** Otros ajustes */
  pageSize: number;

  /** URL para el botón "Go Back" */
  goBackUrl?: string;
}

export default class PORequestFormWebPart
  extends BaseClientSideWebPart<IPORequestFormWebPartProps> {

  private _sp: SPFI;

  /** Inicializa PnPjs con el contexto SPFx */
 protected async onInit(): Promise<void> {
  this._sp = spfi().using(SPFx(this.context));

  if (this.properties.pageSize === undefined) this.properties.pageSize = 25;

  if (!this.properties.parentListId) this.properties.parentListId = null;
  if (!this.properties.childListId) this.properties.childListId = null;
  if (!this.properties.suppliersListId) this.properties.suppliersListId = null;
  if (!this.properties.supplierChildListId) this.properties.supplierChildListId = null;
  if (!this.properties.projectsListId) this.properties.projectsListId = null;
  if (!this.properties.glCodesListId) this.properties.glCodesListId = null;
  // ✅ CAMBIO AQUÍ: usar null en lugar de ''
  if (!this.properties.authorizersListId) this.properties.authorizersListId = null;
  if(!this.properties.authorizersListTitle) this.properties.authorizersListTitle = undefined;
  if (!this.properties.companiesListId) this.properties.companiesListId = null;
  if (!this.properties.approversListId) this.properties.approversListId = null;

  if (this.properties.goBackUrl === undefined) this.properties.goBackUrl = '';

  if (!(window as any).__fluentIconsInitialized) {
    initializeIcons();
    (window as any).__fluentIconsInitialized = true;
  }

  return super.onInit();
}
  public render(): void {
    const element: React.ReactElement<IPoRequestFormProps> = React.createElement(
      PORequestForm,
      {
        context: this.context,

        parentListId: this.properties.parentListId,
        parentListTitle: this.properties.parentListTitle,

        authorizersListId: this.properties.authorizersListId,
        authorizersListTitle: this.properties.authorizersListTitle,
        childListId: this.properties.childListId,
        childListTitle: this.properties.childListTitle,

        suppliersListId: this.properties.suppliersListId,
        suppliersListTitle: this.properties.suppliersListTitle,

        /** ✅ Pasamos también la nueva lista hija de suppliers */
        supplierChildListId: this.properties.supplierChildListId,
        supplierChildListTitle: this.properties.supplierChildListTitle,

        /** ✅ Coding dropdown lists */
        projectsListId: this.properties.projectsListId,
        projectsListTitle: this.properties.projectsListTitle,

        glCodesListId: this.properties.glCodesListId,
        glCodesListTitle: this.properties.glCodesListTitle,

        /** ✅ Company dropdown list */
        companiesListId: this.properties.companiesListId,
        companiesListTitle: this.properties.companiesListTitle,

        /** ✅ Approvers list (NEW) */
        approversListId: this.properties.approversListId,
        approversListTitle: this.properties.approversListTitle,

        pageSize: this.properties.pageSize ?? 25,
        goBackUrl: this.properties.goBackUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'parentListId') {
      this._setListTitleFromId(newValue, 'parentListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    if (propertyPath === 'childListId') {
      this._setListTitleFromId(newValue, 'childListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    if (propertyPath === 'suppliersListId') {
      this._setListTitleFromId(newValue, 'suppliersListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    /** ✅ Nueva lista hija de suppliers */
    if (propertyPath === 'supplierChildListId') {
      this._setListTitleFromId(newValue, 'supplierChildListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    /** ✅ Projects list */
    if (propertyPath === 'projectsListId') {
      this._setListTitleFromId(newValue, 'projectsListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    /** ✅ GL Codes list */
    if (propertyPath === 'glCodesListId') {
      this._setListTitleFromId(newValue, 'glCodesListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    /** ✅ Companies list */
    if (propertyPath === 'companiesListId') {
      this._setListTitleFromId(newValue, 'companiesListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }

    /** ✅ Approvers list (NEW) */
    if (propertyPath === 'approversListId') {
      this._setListTitleFromId(newValue, 'approversListTitle')
        .then(() => this.render())
        .catch(() => this.render());
    }
  }

  /** Busca el título de una lista por GUID y lo guarda en las propiedades */
  private async _setListTitleFromId(
    listId: string,
    titleProp:
      | 'parentListTitle'
      | 'childListTitle'
      | 'suppliersListTitle'
      | 'supplierChildListTitle'
      | 'projectsListTitle'
      | 'glCodesListTitle'
      | 'companiesListTitle'
      | 'approversListTitle'
  ): Promise<void> {
    try {
      if (listId) {
        const li = await this._sp.web.lists.getById(listId)();
        (this.properties as any)[titleProp] = li?.Title;
      } else {
        (this.properties as any)[titleProp] = undefined;
      }
    } catch {
      (this.properties as any)[titleProp] = undefined;
    }
  }

  /** Panel de propiedades: pickers para listas y ajustes */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Conexiones a listas' },
          groups: [
            {
              groupName: 'Listas (Padre / Hijas / Proveedores / Coding / Approvers)',
              groupFields: [
                // Lista PADRE
                PropertyFieldListPicker('parentListId', {
                  label: 'Lista PADRE (PO Requests)',
                  selectedList: this.properties.parentListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'parentListPicker'
                }),
                PropertyPaneTextField('parentListTitle', {
                  label: 'Nombre (solo lectura) - PADRE',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // Lista HIJA (Lines)
                PropertyFieldListPicker('childListId', {
                  label: 'Lista HIJA (PO Request Lines)',
                  selectedList: this.properties.childListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'childListPicker'
                }),
                PropertyPaneTextField('childListTitle', {
                  label: 'Nombre (solo lectura) - HIJA (Lines)',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // Lista de Proveedores sugeridos (antigua)
                PropertyFieldListPicker('suppliersListId', {
                  label: 'Lista de Proveedores sugeridos (antigua)',
                  selectedList: this.properties.suppliersListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'suppliersListPicker'
                }),
                PropertyPaneTextField('suppliersListTitle', {
                  label: 'Nombre (solo lectura) - Proveedores sugeridos',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // ✅ Nueva lista hija de Suppliers (directa al PR)
                PropertyFieldListPicker('supplierChildListId', {
                  label: 'Lista HIJA (Suppliers - nueva)',
                  selectedList: this.properties.supplierChildListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'supplierChildListPicker'
                }),
                PropertyPaneTextField('supplierChildListTitle', {
                  label: 'Nombre (solo lectura) - HIJA (Suppliers - nueva)',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // ✅ Projects (dropdown)
                PropertyFieldListPicker('projectsListId', {
                  label: 'Lista de Projects (dropdown)',
                  selectedList: this.properties.projectsListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'projectsListPicker'
                }),
                PropertyPaneTextField('projectsListTitle', {
                  label: 'Nombre (solo lectura) - Projects',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // ✅ GL Codes (dropdown)
                PropertyFieldListPicker('glCodesListId', {
                  label: 'Lista de GL Codes (dropdown)',
                  selectedList: this.properties.glCodesListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'glCodesListPicker'
                }),
                PropertyPaneTextField('glCodesListTitle', {
                  label: 'Nombre (solo lectura) - GL Codes',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // ✅ Companies (dropdown)
                PropertyFieldListPicker('companiesListId', {
                  label: 'Lista de Companies (dropdown)',
                  selectedList: this.properties.companiesListId || undefined,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 200,
                  key: 'companiesListPicker'
                }),
                PropertyPaneTextField('companiesListTitle', {
                  label: 'Nombre (solo lectura) - Companies',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                }),

                // ✅ Approvers (NEW)
               PropertyFieldListPicker('authorizersListId', {
  label: 'Lista de Authorizers (reglas/aprobadores)',
  selectedList: this.properties.authorizersListId || undefined,  // ← Importante: usar undefined si es null
  includeHidden: false,
  orderBy: PropertyFieldListPickerOrderBy.Title,
  disabled: false,
  multiSelect: false,
  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
  properties: this.properties,
  context: this.context,
  deferredValidationTime: 200,
  key: 'authorizersListPicker'
}),
                PropertyPaneTextField('authorizersListTitle', {
                  label: 'Nombre (solo lectura) - Authorizers',
                  description: 'Se completa automáticamente al elegir la lista',
                  multiline: false,
                  disabled: true
                })
              ]
            },
            {
              groupName: 'Comportamiento',
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  label: 'Tamaño de página (líneas)',
                  min: 5,
                  max: 100,
                  step: 5,
                  value: this.properties.pageSize ?? 25,
                  showValue: true
                }),
                PropertyPaneTextField('goBackUrl', {
                  label: 'Go Back URL',
                  description: 'URL a la que se navega al pulsar "Go Back".',
                  multiline: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
