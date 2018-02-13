

#include <QtGui/QDialog>


#include "ui_SqlQueryDialog.h"


class SqlQueryDialog : public QDialog, private Ui::SqlQueryDialog
{

    Q_OBJECT

public:

    SqlQueryDialog( QWidget * parent = 0 );

public slots:

    void accept( void );

};
