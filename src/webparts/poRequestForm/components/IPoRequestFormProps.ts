import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPoRequestFormProps {
  context: WebPartContext;

  parentListId: string | null;
  parentListTitle?: string;

  childListId: string | null;
  childListTitle?: string;

  suppliersListId: string | null;
  suppliersListTitle?: string;

  supplierChildListId: string | null;
  supplierChildListTitle?: string;

  /** ✅ NUEVAS LISTAS */
  projectsListId: string | null;
  projectsListTitle?: string;

  /** ✅ NUEVA LISTA PARA COMPANY (dropdown) */
  companiesListId: string | null;
  companiesListTitle?: string;

  glCodesListId: string | null;
  glCodesListTitle?: string;
  glCodeListTitle?: string;

  /** ✅ NUEVA LISTA: APPROVERS (reglas / aprobadores) */
  approversListId: string | null;
  approversListTitle?: string;
authorizersListId: string;
  authorizersListTitle?: string;
  pageSize: number;
  goBackUrl?: string;
}

export default IPoRequestFormProps;
