

#include <QtGui/QDialog>


#include "ui_ProfileDialog.h"


class ProfileDialog : public QDialog, private Ui::ProfileDialog
{

    Q_OBJECT

public:

    ProfileDialog( QWidget * parent = 0 );

public slots:

    void accept( void );

};
