

#include "SqlQueryDialog.h"


///
///  @brief  
///
SqlQueryDialog::SqlQueryDialog( QWidget * parent )
  :  QDialog( parent )
{
    setupUi( this );

    return;
}


///
///  @brief  
///
void
SqlQueryDialog::accept( void )
{
    QDialog::accept();

    return;
}
