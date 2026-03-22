#pragma once
#include <QString>
#include <QVariant>
#include <functional>

class Sheet;

class FormulaEngine {
public:
    // Evaluate formula string. sheet is used to resolve cell references.
    static QVariant evaluate(const QString& formula, Sheet* sheet, int baseRow, int baseCol);

    // Format a value according to NumFmt
    static QString formatValue(const QVariant& value, int numFmt, int decimals);

    // Column label helpers
    static QString  colLabel(int col);   // 0→"A", 25→"Z", 26→"AA"
    static int      colIndex(const QString& label); // "A"→0, "AA"→26
    static QString  cellRef(int row, int col);       // (0,0)→"A1"

private:
    using FnArgs = QVector<QVariant>;
    using FnMap  = QHash<QString, std::function<QVariant(FnArgs)>>;

    static FnMap buildFunctions(Sheet* sheet, int r0, int c0);
    static QVariant callFunction(const QString& name, const FnArgs& args,
                                  Sheet* sheet, int r0, int c0);
    // Tokeniser / parser helpers
    static QString  expandRanges(const QString& expr, Sheet* sheet);
    static QString  expandCellRefs(const QString& expr, Sheet* sheet);
    static QVariant evalExpr(const QString& expr, Sheet* sheet, int r0, int c0);
};
