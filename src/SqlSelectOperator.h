#if ! defined(SQLSELECTOPERATOR_H)
#define SQLSELECTOPERATOR_H


#include <QtGui/QComboBox>


class SqlSelectOperator : public QComboBox
{
    Q_OBJECT

public:

    SqlSelectOperator( QWidget * parent = 0 );

private:

    void setSqlOperators( void );

    static const int numSqlOperators = 8;
    static const QString sqlOperatorArr[numSqlOperators];
};

#endif // SQLSELECTOPERATOR_H
