

#include "AnalysisInputDialog.h"

const double AnalysisInputDialog::permittivity[9] = {
  1.0,
  69.18311,
  31.62278,
  81.28304,
  14.79108,
  3.162278,
  3.162278,
  3.162278,
  1.0
};

const double AnalysisInputDialog::conductivity[9] = {
  0.00001,
  5.01187,
  0.1344505,
  0.1312451,
  2.820789e-2,
  1.315351e-4,
  7.713176e-4,
  2.297889e-4,
  0.00001
};

///
///  @brief  
///
AnalysisInputDialog::AnalysisInputDialog( QWidget * parent )
  :  QDialog( parent )
{
    setupUi( this );

    connect( comboBoxGroundType, SIGNAL(currentIndexChanged(int)),
             this, SLOT(setGroundParams(int)) );

    return;
}


///
///  @brief  
///
void
AnalysisInputDialog::setGroundParams( int index )
{
    if ( index == 0 )
    {
        doubleSpinBoxPermittivity->setEnabled( true );
        doubleSpinBoxConductivity->setEnabled( true );
    }
    else
    {
        doubleSpinBoxPermittivity->setEnabled( false );
        doubleSpinBoxConductivity->setEnabled( false );
        doubleSpinBoxPermittivity->setValue( permittivity[index] );
        doubleSpinBoxConductivity->setValue( conductivity[index] );
    }

    return;
}


///
///  @brief  
///
void
AnalysisInputDialog::accept( void )
{
    QDialog::accept();

    return;
}
