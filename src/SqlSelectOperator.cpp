

#include <QtCore/QTextStream>
#include <QtGui/QColorDialog>


#include "SqlSelectOperator.h"


const QString SqlSelectOperator::sqlOperatorArr[numSqlOperators] =
{
  tr("All"),
  "<",
  "=",
  ">",
  "<=",
  ">=",
  "<>",
  "between"
};


///
///  @brief  
///
SqlSelectOperator::SqlSelectOperator( QWidget * parent )
 : QComboBox( parent )
{
    setSqlOperators();

    return;
}


///
///  @brief  
///
void
SqlSelectOperator::setSqlOperators( void )
{
    for ( int o = 0; o < numSqlOperators; ++o )
    {
        addItem( sqlOperatorArr[o] );
    }

    return;
}
