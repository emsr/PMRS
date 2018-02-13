

#include "SqlQueryResultsDialog.h"


///
///  @brief  
///
SqlQueryResultsDialog::SqlQueryResultsDialog( QWidget * parent )
  :  QDialog( parent )
{
    setupUi( this );

    return;
}


///
///  @brief  
///
void
SqlQueryResultsDialog::accept( void )
{
    QDialog::accept();

    return;
}
