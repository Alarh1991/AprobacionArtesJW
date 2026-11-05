import * as React from 'react';
import styles from './FormularioArtes.module.scss';
import type { IFormularioArtesProps } from './IFormularioArtesProps';

import { ApprovalForm } from './approval-form/ApprovalForm';

interface IErrorBoundaryState {
  hasError: boolean;
  errorMessage: string;
}

class ApprovalFormErrorBoundary extends React.Component<{}, IErrorBoundaryState> {
  public state: IErrorBoundaryState = {
    hasError: false,
    errorMessage: ''
  };

  public componentDidCatch(error: Error): void {
    this.setState({
      hasError: true,
      errorMessage: error?.message ?? 'Error desconocido al renderizar el formulario.'
    });
    // Deja rastro en consola para depuración en el navegador
    // eslint-disable-next-line no-console
    console.error('Error renderizando ApprovalForm:', error);
  }

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '12px', border: '1px solid #d13438', color: '#d13438' }}>
          Error de render en el web part: {this.state.errorMessage}
        </div>
      );
    }

    return this.props.children;
  }
}

export default class FormularioArtes extends React.Component<IFormularioArtesProps> {
  public render(): React.ReactElement<IFormularioArtesProps> {
    return (
      <section className={styles.formularioArtes}>
        <ApprovalFormErrorBoundary>
          <ApprovalForm />
        </ApprovalFormErrorBoundary>
      </section>
    );
  }
}
