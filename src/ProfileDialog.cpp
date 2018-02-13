

#include "ProfileDialog.h"


///
///  @brief  
///
ProfileDialog::ProfileDialog( QWidget * parent )
  :  QDialog( parent )
{
    setupUi( this );

    return;
}


///
///  @brief  
///
void
ProfileDialog::accept( void )
{
    QDialog::accept();

    return;
}
