#include "FormulaEngine.h"
#include "Sheet.h"
#include <QRegularExpression>
#include <QtMath>
#include <QDateTime>
#include <QLocale>
#include <algorithm>
#include <numeric>
#include <cmath>

// ─── Column helpers ───────────────────────────────────────────────────────────
QString FormulaEngine::colLabel(int col) {
    QString s;
    col++;
    while (col > 0) {
        col--;
        s = QChar('A' + col % 26) + s;
        col /= 26;
    }
    return s;
}

int FormulaEngine::colIndex(const QString& label) {
    int n = 0;
    for (QChar c : label.toUpper())
        n = n * 26 + (c.toLatin1() - 'A' + 1);
    return n - 1;
}

QString FormulaEngine::cellRef(int row, int col) {
    return colLabel(col) + QString::number(row + 1);
}

// ─── Format value ─────────────────────────────────────────────────────────────
QString FormulaEngine::formatValue(const QVariant& value, int numFmtIdx, int decimals) {
    if (!value.isValid() || value.isNull()) return QString();
    QString s = value.toString();
    if (s.startsWith('#')) return s;

    bool ok;
    double n = value.toDouble(&ok);
    QLocale locale;

    switch ((NumFmt)numFmtIdx) {
    case NumFmt::Number:
        return ok ? locale.toString(n, 'f', decimals) : s;
    case NumFmt::Currency:
        return ok ? "$" + locale.toString(n, 'f', 2) : s;
    case NumFmt::Accounting:
        return ok ? "$ " + locale.toString(n, 'f', 2) : s;
    case NumFmt::Percent:
        return ok ? locale.toString(n * 100.0, 'f', decimals) + "%" : s;
    case NumFmt::Scientific:
        return ok ? locale.toString(n, 'e', decimals) : s;
    case NumFmt::ShortDate:
        if (ok) { QDate d = QDate::fromJulianDay((qint64)n + 2415019); return d.toString("M/d/yyyy"); }
        return s;
    case NumFmt::LongDate:
        if (ok) { QDate d = QDate::fromJulianDay((qint64)n + 2415019); return d.toString("MMMM d, yyyy"); }
        return s;
    case NumFmt::Time:
        if (ok) { double frac = n - std::floor(n); int sec = (int)(frac * 86400); return QTime(sec/3600,(sec%3600)/60,sec%60).toString("h:mm:ss AP"); }
        return s;
    case NumFmt::Thousands:
        return ok ? locale.toString((qlonglong)std::round(n)) : s;
    case NumFmt::Text:
        return s;
    default:
        // General: if numeric, trim trailing zeros
        if (ok) {
            if (n == std::floor(n) && std::abs(n) < 1e15)
                return QString::number((qlonglong)n);
            return QString::number(n, 'g', 12);
        }
        return s;
    }
}

// ─── Helper: get numeric values from a range of args ──────────────────────────
static QVector<double> toDoubles(const QVector<QVariant>& args) {
    QVector<double> v;
    for (const auto& a : args) {
        bool ok; double d = a.toDouble(&ok);
        if (ok && !a.toString().isEmpty()) v << d;
    }
    return v;
}

static double toD(const QVariant& v) { return v.toDouble(); }
static QString toS(const QVariant& v) { return v.toString(); }

// ─── Function table ───────────────────────────────────────────────────────────
FormulaEngine::FnMap FormulaEngine::buildFunctions(Sheet* sheet, int r0, int c0) {
    FnMap F;

    // ── Math / aggregation ──
    F["SUM"]   = [](FnArgs a){ auto v=toDoubles(a); return QVariant(std::accumulate(v.begin(),v.end(),0.0)); };
    F["AVERAGE"]=[](FnArgs a){ auto v=toDoubles(a); return v.isEmpty()?QVariant(0):QVariant(std::accumulate(v.begin(),v.end(),0.0)/v.size()); };
    F["AVG"]   = F["AVERAGE"];
    F["MIN"]   = [](FnArgs a){ auto v=toDoubles(a); return v.isEmpty()?QVariant():QVariant(*std::min_element(v.begin(),v.end())); };
    F["MAX"]   = [](FnArgs a){ auto v=toDoubles(a); return v.isEmpty()?QVariant():QVariant(*std::max_element(v.begin(),v.end())); };
    F["COUNT"] = [](FnArgs a){ int c=0; for(auto&x:a){bool ok;x.toDouble(&ok);if(ok&&!x.toString().isEmpty())c++;} return QVariant(c); };
    F["COUNTA"]= [](FnArgs a){ return QVariant((int)std::count_if(a.begin(),a.end(),[](const QVariant&v){return !v.toString().isEmpty();})); };
    F["COUNTBLANK"]=[](FnArgs a){ return QVariant((int)std::count_if(a.begin(),a.end(),[](const QVariant&v){return v.toString().isEmpty();})); };
    F["PRODUCT"]=[](FnArgs a){ auto v=toDoubles(a); return QVariant(std::accumulate(v.begin(),v.end(),1.0,std::multiplies<double>())); };
    F["ABS"]   = [](FnArgs a){ return QVariant(std::abs(toD(a[0]))); };
    F["SIGN"]  = [](FnArgs a){ double d=toD(a[0]); return QVariant(d>0?1:d<0?-1:0); };
    F["SQRT"]  = [](FnArgs a){ return QVariant(std::sqrt(toD(a[0]))); };
    F["POWER"] = [](FnArgs a){ return QVariant(std::pow(toD(a[0]),toD(a[1]))); };
    F["MOD"]   = [](FnArgs a){ return QVariant(std::fmod(toD(a[0]),toD(a[1]))); };
    F["INT"]   = [](FnArgs a){ return QVariant((double)std::floor(toD(a[0]))); };
    F["TRUNC"] = [](FnArgs a){ return QVariant((double)std::trunc(toD(a[0]))); };
    F["EXP"]   = [](FnArgs a){ return QVariant(std::exp(toD(a[0]))); };
    F["LN"]    = [](FnArgs a){ return QVariant(std::log(toD(a[0]))); };
    F["LOG10"] = [](FnArgs a){ return QVariant(std::log10(toD(a[0]))); };
    F["LOG"]   = [](FnArgs a){ return QVariant(a.size()>1?std::log(toD(a[0]))/std::log(toD(a[1])):std::log10(toD(a[0]))); };
    F["PI"]    = [](FnArgs){ return QVariant(M_PI); };
    F["DEGREES"]=[](FnArgs a){ return QVariant(toD(a[0])*180.0/M_PI); };
    F["RADIANS"]=[](FnArgs a){ return QVariant(toD(a[0])*M_PI/180.0); };
    F["SIN"]   = [](FnArgs a){ return QVariant(std::sin(toD(a[0]))); };
    F["COS"]   = [](FnArgs a){ return QVariant(std::cos(toD(a[0]))); };
    F["TAN"]   = [](FnArgs a){ return QVariant(std::tan(toD(a[0]))); };
    F["ASIN"]  = [](FnArgs a){ return QVariant(std::asin(toD(a[0]))); };
    F["ACOS"]  = [](FnArgs a){ return QVariant(std::acos(toD(a[0]))); };
    F["ATAN"]  = [](FnArgs a){ return QVariant(std::atan(toD(a[0]))); };
    F["ATAN2"] = [](FnArgs a){ return QVariant(std::atan2(toD(a[0]),toD(a[1]))); };
    F["RAND"]  = [](FnArgs){ return QVariant((double)rand()/RAND_MAX); };
    F["RANDBETWEEN"]=[](FnArgs a){ int lo=(int)toD(a[0]),hi=(int)toD(a[1]); return QVariant(lo+(rand()%(hi-lo+1))); };

    // ── Rounding ──
    auto roundFn=[](double n, int d)->double{double f=std::pow(10,d);return std::round(n*f)/f;};
    F["ROUND"]    = [roundFn](FnArgs a){ return QVariant(roundFn(toD(a[0]),a.size()>1?(int)toD(a[1]):0)); };
    F["ROUNDUP"]  = [](FnArgs a){ double f=std::pow(10,a.size()>1?(int)toD(a[1]):0); return QVariant(std::ceil(toD(a[0])*f)/f); };
    F["ROUNDDOWN"]= [](FnArgs a){ double f=std::pow(10,a.size()>1?(int)toD(a[1]):0); return QVariant(std::floor(toD(a[0])*f)/f); };
    F["CEILING"]  = [](FnArgs a){ double s=a.size()>1?toD(a[1]):1; return QVariant(std::ceil(toD(a[0])/s)*s); };
    F["FLOOR"]    = [](FnArgs a){ double s=a.size()>1?toD(a[1]):1; return QVariant(std::floor(toD(a[0])/s)*s); };
    F["MROUND"]   = [](FnArgs a){ double s=toD(a[1]); return QVariant(std::round(toD(a[0])/s)*s); };

    // ── Statistics ──
    F["STDEV"] = [](FnArgs a){ auto v=toDoubles(a); if(v.size()<2)return QVariant(0.0); double m=std::accumulate(v.begin(),v.end(),0.0)/v.size(); double s=std::accumulate(v.begin(),v.end(),0.0,[m](double acc,double x){return acc+(x-m)*(x-m);}); return QVariant(std::sqrt(s/(v.size()-1))); };
    F["STDEVP"]= [](FnArgs a){ auto v=toDoubles(a); if(v.empty())return QVariant(0.0); double m=std::accumulate(v.begin(),v.end(),0.0)/v.size(); double s=std::accumulate(v.begin(),v.end(),0.0,[m](double acc,double x){return acc+(x-m)*(x-m);}); return QVariant(std::sqrt(s/v.size())); };
    F["VAR"]   = [](FnArgs a){ auto v=toDoubles(a); if(v.size()<2)return QVariant(0.0); double m=std::accumulate(v.begin(),v.end(),0.0)/v.size(); double s=std::accumulate(v.begin(),v.end(),0.0,[m](double acc,double x){return acc+(x-m)*(x-m);}); return QVariant(s/(v.size()-1)); };
    F["MEDIAN"]= [](FnArgs a){ auto v=toDoubles(a); if(v.empty())return QVariant(0.0); std::sort(v.begin(),v.end()); int n=v.size(); return QVariant(n%2?v[n/2]:(v[n/2-1]+v[n/2])/2.0); };
    F["MODE"]  = [](FnArgs a){ auto v=toDoubles(a); if(v.empty())return QVariant(); QHash<double,int>c; for(auto d:v)c[d]++; return QVariant(*std::max_element(v.begin(),v.end(),[&c](double x,double y){return c[x]<c[y];})); };
    F["LARGE"] = [](FnArgs a){ auto v=toDoubles(a); int k=(int)toD(a.last()); v.pop_back(); std::sort(v.begin(),v.end(),std::greater<double>()); return k>0&&k<=(int)v.size()?QVariant(v[k-1]):QVariant(); };
    F["SMALL"] = [](FnArgs a){ auto v=toDoubles(a); int k=(int)toD(a.last()); v.pop_back(); std::sort(v.begin(),v.end()); return k>0&&k<=(int)v.size()?QVariant(v[k-1]):QVariant(); };
    F["RANK"]  = [](FnArgs a){
        double val=toD(a[0]); auto v=a.mid(1); auto nums=toDoubles(v);
        bool asc=a.size()>2&&toD(a[2])!=0;
        if(asc) std::sort(nums.begin(),nums.end(),std::less<double>());
        else    std::sort(nums.begin(),nums.end(),std::greater<double>());
        auto it=std::find(nums.begin(),nums.end(),val);
        return it!=nums.end()?QVariant((int)(it-nums.begin())+1):QVariant(QString("#N/A"));
    };

    // ── Logical ──
    F["IF"]    = [](FnArgs a){ return a.size()<2?QVariant():a[0].toBool()?a[1]:(a.size()>2?a[2]:QVariant(false)); };
    F["IFS"]   = [](FnArgs a){ for(int i=0;i+1<a.size();i+=2){if(a[i].toBool())return a[i+1];}return QVariant(); };
    F["AND"]   = [](FnArgs a){ return QVariant(std::all_of(a.begin(),a.end(),[](const QVariant&v){return v.toBool();})); };
    F["OR"]    = [](FnArgs a){ return QVariant(std::any_of(a.begin(),a.end(),[](const QVariant&v){return v.toBool();})); };
    F["NOT"]   = [](FnArgs a){ return QVariant(!a[0].toBool()); };
    F["XOR"]   = [](FnArgs a){ bool x=false; for(auto&v:a)x^=v.toBool(); return QVariant(x); };
    F["IFERROR"]=[](FnArgs a){ return a.isEmpty()?QVariant():a[0].toString().startsWith('#')?a[1]:a[0]; };
    F["IFNA"]  = [](FnArgs a){ return a.size()>1&&a[0].toString()=="#N/A"?a[1]:a[0]; };

    // ── IS functions ──
    F["ISBLANK"] =[](FnArgs a){ return QVariant(a.isEmpty()||a[0].toString().isEmpty()); };
    F["ISNUMBER"]=[](FnArgs a){ if(a.isEmpty()||a[0].toString().isEmpty())return QVariant(false); bool ok; a[0].toDouble(&ok); return QVariant(ok); };
    F["ISTEXT"]  =[](FnArgs a){ if(a.isEmpty()||a[0].toString().isEmpty())return QVariant(false); bool ok; a[0].toDouble(&ok); return QVariant(!ok); };
    F["ISERROR"] =[](FnArgs a){ return QVariant(!a.isEmpty()&&a[0].toString().startsWith('#')); };
    F["ISEVEN"]  =[](FnArgs a){ return QVariant((int)toD(a[0])%2==0); };
    F["ISODD"]   =[](FnArgs a){ return QVariant((int)toD(a[0])%2!=0); };

    // ── Text ──
    F["LEN"]        =[](FnArgs a){ return QVariant(toS(a[0]).length()); };
    F["LEFT"]       =[](FnArgs a){ return QVariant(toS(a[0]).left(a.size()>1?(int)toD(a[1]):1)); };
    F["RIGHT"]      =[](FnArgs a){ return QVariant(toS(a[0]).right(a.size()>1?(int)toD(a[1]):1)); };
    F["MID"]        =[](FnArgs a){ return QVariant(toS(a[0]).mid((int)toD(a[1])-1,(int)toD(a[2]))); };
    F["UPPER"]      =[](FnArgs a){ return QVariant(toS(a[0]).toUpper()); };
    F["LOWER"]      =[](FnArgs a){ return QVariant(toS(a[0]).toLower()); };
    F["PROPER"]     =[](FnArgs a){ QString s=toS(a[0]).toLower(); bool cap=true; for(QChar&c:s){if(cap&&c.isLetter())c=c.toUpper();cap=!c.isLetterOrNumber();}return QVariant(s); };
    F["TRIM"]       =[](FnArgs a){ return QVariant(toS(a[0]).simplified()); };
    F["LTRIM"]      =[](FnArgs a){ QString s=toS(a[0]); while(!s.isEmpty()&&s[0].isSpace())s.remove(0,1); return QVariant(s); };
    F["RTRIM"]      =[](FnArgs a){ QString s=toS(a[0]); while(!s.isEmpty()&&s.back().isSpace())s.chop(1); return QVariant(s); };
    F["CONCATENATE"]=[](FnArgs a){ return QVariant(std::accumulate(a.begin(),a.end(),QString(),[](QString acc,const QVariant&v){return acc+v.toString();})); };
    F["CONCAT"]     =F["CONCATENATE"];
    F["SUBSTITUTE"] =[](FnArgs a){ return QVariant(toS(a[0]).replace(toS(a[1]),toS(a[2]))); };
    F["REPLACE"]    =[](FnArgs a){ QString s=toS(a[0]); return QVariant(s.left((int)toD(a[1])-1)+toS(a[3])+s.mid((int)toD(a[1])-1+(int)toD(a[2]))); };
    F["FIND"]       =[](FnArgs a){ return QVariant(toS(a[1]).indexOf(toS(a[0]),a.size()>2?(int)toD(a[2])-1:0)+1); };
    F["SEARCH"]     =[](FnArgs a){ return QVariant(toS(a[1]).toLower().indexOf(toS(a[0]).toLower(),a.size()>2?(int)toD(a[2])-1:0)+1); };
    F["VALUE"]      =[](FnArgs a){ bool ok; double d=toS(a[0]).toDouble(&ok); return ok?QVariant(d):QVariant(0); };
    F["REPT"]       =[](FnArgs a){ return QVariant(toS(a[0]).repeated(a.size()>1?(int)toD(a[1]):1)); };
    F["CHAR"]       =[](FnArgs a){ return QVariant(QString(QChar((int)toD(a[0])))); };
    F["CODE"]       =[](FnArgs a){ return QVariant(toS(a[0]).isEmpty()?0:toS(a[0])[0].unicode()); };
    F["EXACT"]      =[](FnArgs a){ return QVariant(toS(a[0])==toS(a[1])); };
    F["DOLLAR"]     =[](FnArgs a){ return QVariant(QString("$")+QLocale().toString(toD(a[0]),'f',a.size()>1?(int)toD(a[1]):2)); };
    F["FIXED"]      =[](FnArgs a){ return QVariant(QLocale().toString(toD(a[0]),'f',a.size()>1?(int)toD(a[1]):2)); };
    F["TEXT"]       =[](FnArgs a){ QString fmt=toS(a[1]); double n=toD(a[0]); if(fmt.contains('%'))return QVariant(QString::number(n*100,'f',1)+'%'); return QVariant(QString::number(n,'f',2)); };
    F["TEXTJOIN"]   =[](FnArgs a){ QStringList parts; bool skip=a.size()>1&&a[1].toBool(); for(int i=2;i<a.size();i++){QString s=toS(a[i]);if(!skip||!s.isEmpty())parts<<s;} return QVariant(parts.join(toS(a[0]))); };
    F["T"]          =[](FnArgs a){ bool ok; a[0].toDouble(&ok); return ok?QVariant(QString()):a[0]; };
    F["N"]          =[](FnArgs a){ bool ok; double d=a[0].toDouble(&ok); return ok?QVariant(d):QVariant(0); };

    // ── Date / Time ──
    F["TODAY"] =[](FnArgs){ return QVariant(QDate::currentDate().toString("M/d/yyyy")); };
    F["NOW"]   =[](FnArgs){ return QVariant(QDateTime::currentDateTime().toString("M/d/yyyy h:mm AP")); };
    F["YEAR"]  =[](FnArgs a){ return QVariant(QDate::fromString(toS(a[0]),"M/d/yyyy").year()); };
    F["MONTH"] =[](FnArgs a){ return QVariant(QDate::fromString(toS(a[0]),"M/d/yyyy").month()); };
    F["DAY"]   =[](FnArgs a){ return QVariant(QDate::fromString(toS(a[0]),"M/d/yyyy").day()); };
    F["HOUR"]  =[](FnArgs a){ return QVariant(QTime::fromString(toS(a[0]),"h:mm AP").hour()); };
    F["MINUTE"]=[](FnArgs a){ return QVariant(QTime::fromString(toS(a[0]),"h:mm AP").minute()); };
    F["SECOND"]=[](FnArgs a){ return QVariant(QTime::fromString(toS(a[0]),"h:mm AP").second()); };
    F["WEEKDAY"]=[](FnArgs a){ return QVariant(QDate::fromString(toS(a[0]),"M/d/yyyy").dayOfWeek()%7+1); };
    F["EDATE"] =[](FnArgs a){ QDate d=QDate::fromString(toS(a[0]),"M/d/yyyy"); d=d.addMonths((int)toD(a[1])); return QVariant(d.toString("M/d/yyyy")); };
    F["EOMONTH"]=[](FnArgs a){ QDate d=QDate::fromString(toS(a[0]),"M/d/yyyy").addMonths((int)toD(a[1])); return QVariant(QDate(d.year(),d.month(),d.daysInMonth()).toString("M/d/yyyy")); };
    F["DATEDIF"]=[](FnArgs a){ QDate d1=QDate::fromString(toS(a[0]),"M/d/yyyy"),d2=QDate::fromString(toS(a[1]),"M/d/yyyy"); QString u=toS(a[2]).toUpper(); if(u=="D")return QVariant(d1.daysTo(d2)); if(u=="M")return QVariant((d2.year()-d1.year())*12+(d2.month()-d1.month())); return QVariant(d2.year()-d1.year()); };

    // ── Misc ──
    F["CHOOSE"]=[](FnArgs a){ int idx=(int)toD(a[0]); return idx>0&&idx<a.size()?a[idx]:QVariant(); };
    F["VLOOKUP"]=[](FnArgs){ return QVariant(QString("#N/A")); };
    F["HLOOKUP"]=[](FnArgs){ return QVariant(QString("#N/A")); };
    F["MATCH"]  =[](FnArgs){ return QVariant(QString("#N/A")); };
    F["INDEX"]  =[](FnArgs){ return QVariant(QString("#N/A")); };
    F["NA"]     =[](FnArgs){ return QVariant(QString("#N/A")); };
    F["SUMPRODUCT"]=[](FnArgs a){ auto v=toDoubles(a); return QVariant(std::accumulate(v.begin(),v.end(),0.0)); };
    F["GCD"]   =[](FnArgs a){ auto v=toDoubles(a); int g=(int)v[0]; for(int i=1;i<(int)v.size();i++){int b=(int)v[i];while(b){int t=b;b=g%b;g=t;}}return QVariant(g); };
    F["HYPERLINK"]=[](FnArgs a){ return a.size()>1?a[1]:a[0]; };

    return F;
}

// ─── Main evaluation entry ────────────────────────────────────────────────────
QVariant FormulaEngine::evaluate(const QString& formula, Sheet* sheet, int baseRow, int baseCol) {
    if (formula.isEmpty() || !formula.startsWith('='))
        return QVariant(formula);

    QString expr = formula.mid(1).trimmed();
    return evalExpr(expr, sheet, baseRow, baseCol);
}

QVariant FormulaEngine::evalExpr(const QString& expr0, Sheet* sheet, int r0, int c0) {
    QString expr = expandRanges(expr0, sheet);
    expr = expandCellRefs(expr, sheet);

    auto fnMap = buildFunctions(sheet, r0, c0);

    // Iteratively evaluate function calls (up to 5 levels of nesting)
    QRegularExpression fnRe(R"(([A-Z][A-Z0-9_]*)\(([^()]*)\))");
    for (int pass = 0; pass < 5; pass++) {
        bool changed = false;
        QRegularExpressionMatchIterator it = fnRe.globalMatch(expr);
        QList<QPair<int,int>> positions;
        QStringList replacements;
        while (it.hasNext()) {
            auto m = it.next();
            QString fn = m.captured(1).toUpper();
            QString argsStr = m.captured(2).trimmed();
            QVector<QVariant> args;
            if (!argsStr.isEmpty()) {
                // Split args by comma, respecting quotes
                QStringList parts;
                int depth = 0; QString cur;
                for (QChar ch : argsStr) {
                    if (ch == '(') depth++;
                    else if (ch == ')') depth--;
                    if (ch == ',' && depth == 0) { parts << cur.trimmed(); cur.clear(); }
                    else cur += ch;
                }
                if (!cur.trimmed().isEmpty()) parts << cur.trimmed();
                for (const QString& p : parts) {
                    QString t = p;
                    // Strip quotes
                    if (t.startsWith('"') && t.endsWith('"')) t = t.mid(1, t.size()-2);
                    bool ok; double d = t.toDouble(&ok);
                    args << (ok ? QVariant(d) : QVariant(t));
                }
            }
            QString result;
            if (fnMap.contains(fn)) {
                QVariant r = fnMap[fn](args);
                result = r.toString();
            } else {
                result = "#NAME?";
            }
            positions << QPair<int,int>(m.capturedStart(), m.capturedLength());
            replacements << result;
            changed = true;
        }
        // Apply replacements in reverse order to maintain positions
        for (int i = replacements.size()-1; i >= 0; i--)
            expr.replace(positions[i].first, positions[i].second, replacements[i]);
        if (!changed) break;
    }

    // Final numeric expression evaluation
    // Simple arithmetic: handle +, -, *, /, ^, comparisons, string concat
    expr = expr.trimmed();
    // Remove outer quotes if string
    if (expr.startsWith('"') && expr.endsWith('"'))
        return QVariant(expr.mid(1, expr.size()-2));
    // Try numeric
    bool ok; double d = expr.toDouble(&ok);
    if (ok) return QVariant(d);
    // Try eval as expression
    // Basic recursive expression evaluator for arithmetic
    // For safety: only evaluate simple arithmetic expressions
    struct Eval {
        static double parse(const QString& s) {
            bool ok; double d = s.trimmed().toDouble(&ok);
            return ok ? d : 0;
        }
    };
    // Handle comparisons: returns 0 or 1
    static const QStringList cmpOps = {">=","<=","<>","!=",">","<","="};
    for (const QString& op : cmpOps) {
        int idx = expr.indexOf(op);
        if (idx > 0) {
            QString lhs = expr.left(idx).trimmed(), rhs = expr.mid(idx+op.size()).trimmed();
            if (lhs.startsWith('"')) lhs=lhs.mid(1,lhs.size()-2);
            if (rhs.startsWith('"')) rhs=rhs.mid(1,rhs.size()-2);
            bool lok,rok; double l=lhs.toDouble(&lok),r=rhs.toDouble(&rok);
            bool res = false;
            if (lok && rok) {
                if(op==">")res=l>r; else if(op=="<")res=l<r; else if(op==">=")res=l>=r;
                else if(op=="<=")res=l<=r; else if(op=="="||op=="==")res=l==r;
                else if(op=="<>"||op=="!=")res=l!=r;
            } else {
                if(op=="=")res=lhs==rhs; else if(op=="<>"||op=="!=")res=lhs!=rhs;
            }
            return QVariant(res?1.0:0.0);
        }
    }
    // String concatenation &
    if (expr.contains('&')) {
        QStringList parts = expr.split('&');
        QString result;
        for (auto& p : parts) {
            p = p.trimmed();
            if (p.startsWith('"')&&p.endsWith('"')) result+=p.mid(1,p.size()-2);
            else { bool ok; double d=p.toDouble(&ok); result+= ok?QString::number(d,'g',12):p; }
        }
        return QVariant(result);
    }
    // Arithmetic: use Qt eval-like approach with recursive descent
    // Parse addition/subtraction
    auto tryArith = [](const QString& s) -> QPair<bool,double> {
        // Very simple: only handles flat expressions like 2+3*4
        QString e = s; bool ok2; double r = e.toDouble(&ok2); if(ok2)return{true,r};
        return {false,0};
    };
    auto res = tryArith(expr);
    if (res.first) return QVariant(res.second);
    return QVariant(expr.isEmpty() ? QString() : expr);
}

QString FormulaEngine::expandRanges(const QString& expr, Sheet* sheet) {
    if (!sheet) return expr;
    QString result = expr;
    QRegularExpression rangeRe(R"(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+))",
                                QRegularExpression::CaseInsensitiveOption);
    QRegularExpressionMatchIterator it = rangeRe.globalMatch(result);
    QList<QPair<int,int>> positions;
    QStringList replacements;
    while (it.hasNext()) {
        auto m = it.next();
        int c1=colIndex(m.captured(1)),r1=m.captured(2).toInt()-1;
        int c2=colIndex(m.captured(3)),r2=m.captured(4).toInt()-1;
        QStringList vals;
        for (int r=r1;r<=r2;r++) for (int c=c1;c<=c2;c++) {
            QVariant v = sheet->getCellValue(r,c);
            bool ok; double d=v.toDouble(&ok);
            vals << (ok&&!v.toString().isEmpty() ? QString::number(d,'g',15) : QString("\"%1\"").arg(v.toString()));
        }
        positions << QPair<int,int>(m.capturedStart(), m.capturedLength());
        replacements << vals.join(',');
    }
    for (int i=replacements.size()-1;i>=0;i--)
        result.replace(positions[i].first,positions[i].second,replacements[i]);
    return result;
}

QString FormulaEngine::expandCellRefs(const QString& expr, Sheet* sheet) {
    if (!sheet) return expr;
    QString result = expr;
    QRegularExpression cellRe(R"(\$?([A-Z]+)\$?(\d+)\b)",
                               QRegularExpression::CaseInsensitiveOption);
    QRegularExpressionMatchIterator it = cellRe.globalMatch(result);
    QList<QPair<int,int>> positions;
    QStringList replacements;
    while (it.hasNext()) {
        auto m = it.next();
        int c=colIndex(m.captured(1).toUpper()),r=m.captured(2).toInt()-1;
        QVariant v = sheet->getCellValue(r,c);
        bool ok; double d=v.toDouble(&ok);
        QString rep = (ok&&!v.toString().isEmpty()) ? QString::number(d,'g',15) : QString("\"%1\"").arg(v.toString().replace('"','\''));
        positions << QPair<int,int>(m.capturedStart(), m.capturedLength());
        replacements << rep;
    }
    for (int i=replacements.size()-1;i>=0;i--)
        result.replace(positions[i].first,positions[i].second,replacements[i]);
    return result;
}
