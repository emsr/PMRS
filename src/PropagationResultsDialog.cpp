

#include "PropagationResultsDialog.h"


///
///  @brief  
///
PropagationResultsDialog::PropagationResultsDialog( QWidget * parent )
  :  QDialog( parent )
{
    setupUi( this );

    return;
}


///
///  @brief  
///
void
PropagationResultsDialog::accept( void )
{
    QDialog::accept();

    return;
}
