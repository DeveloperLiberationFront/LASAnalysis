compatibility_formulas = ['BETADIST', 'BETAINV', 'BINOMDIST', 'CHIDIST', 'CHIINV', 'CHITEST',
                          'CONFIDENCE', 'COVAR', 'CRITBINOM', 'EXPONDIST', 'FDIST',
                          'FINV', 'FLOOR', 'FTEST', 'GAMMADIST', 'GAMMAINV', 'HYPGEOMDIST',
                          'LOGINV', 'LOGNORMDIST', 'MODE', 'NEGBINOMDIST', 'NORMDIST',
                          'NORMINV', 'NORMSDIST', 'NORMSINV', 'PERCENTILE', 'PERCENTRANK',
                          'POISSON', 'QUARTILE', 'RANK', 'STDEV', 'STDEVP', 'TDIST',
                          'TINV', 'TTEST', 'VAR', 'VARP', 'WEIBULL', 'ZTEST']

cube_formulas = ['CUBEKPIMEMBER', 'CUBEMEMBER', 'CUBEMEMBERPROPERTY',
                 'CUBERANKEDMEMBER', 'CUBESET', 'CUBESETCOUNT', 'CUBEVALUE']

database_formulas = ['DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN',
                     'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP']

datetime_formulas = ['DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360',
                     'EDATE', 'EOMONTH', 'HOUR', 'ISOWEEKNUM', 'MINUTE', 'MONTH',
                     'NETWORKDAYS', 'NETWORKDAYS.INTL', 'NOW', 'SECOND', 'TIME',
                     'TIMEVALUE', 'TODAY', 'WEEKDAY', 'WEEKNUM', 'WORKDAY',
                     'WORKDAY.INTL', 'YEAR', 'YEARFRAC']

engineering_formulas = ['BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY', 'BIN2DEC',
                        'BIN2HEX', 'BIN2OCT', 'BITAND', 'BITLSHIFT', 'BITOR',
                        'BITRSHIFT', 'BITXOR', 'COMPLEX', 'CONVERT', 'DEC2BIN',
                        'DEC2HEX', 'DEC2OCT', 'DELTA', 'ERF', 'ERF.PRECISE',
                        'ERFC', 'ERFC.PRECISE', 'GESTEP', 'HEX2BIN', 'HEX2DEC',
                        'HEX2OCT', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE',
                        'IMCOS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMDIV',
                        'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMPOWER',
                        'IMPRODUCT', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN',
                        'IMSINH', 'IMSQRT', 'IMSUB', 'IMSUM', 'IMTAN', 'OCT2BIN',
                        'OCT2DEC', 'OCT2HEX']

financial_formulas = ['ACCRINT', 'ACCRINTM', 'AMORDEGRC', 'AMORLINC', 'COUPDAYBS',
                      'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD',
                      'CUMIPMT', 'CUMPRINC', 'DB', 'DDB', 'DISC', 'DOLLARDE',
                      'DOLLARFR', 'DURATION', 'EFFECT', 'FV', 'FVSCHEDULE',
                      'INTRATE', 'IPMT', 'IRR', 'ISPMT', 'MDURATION', 'MIRR',
                      'NOMINAL', 'NPER', 'NPV', 'ODDFPRICE', 'ODDFYIELD',
                      'ODDLPRICE', 'ODDLYIELD', 'PDURATION', 'PMT', 'PPMT',
                      'PRICE', 'PRICEDISC', 'PRICEMAT', 'PV', 'RATE', 'RECEIVED',
                      'RRI', 'SLN', 'SYD', 'TBILLEQ', 'TBILLPRICE', 'TBILLYIELD',
                      'VDB', 'XIRR', 'XNPV', 'YIELD', 'YIELDDISC', 'YIELDMAT']

information_formulas = ['CELL', 'ERROR.TYPE', 'INFO', 'ISBLANK', 'ISERR', 'ISERROR',
                        'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT',
                        'ISNUMBER', 'ISODD', 'ISREF', 'ISTEXT', 'N', 'NA',
                        'SHEET', 'SHEETS', 'TYPE']

logical_formulas = ['AND', 'FALSE', 'IF', 'IFERROR', 'IFNA', 'NOT', 'OR', 'TRUE',
                    'XOR']

lookup_formulas = ['ADDRESS', 'AREAS', 'CHOOSE', 'COLUMN', 'COLUMNS', 'FORMULATEXT',
                   'GETPIVOTDATA', 'HLOOKUP', 'HYPERLINK', 'INDEX', 'INDIRECT',
                   'LOOKUP', 'MATCH', 'OFFSET', 'ROW', 'ROWS', 'RTD', 'TRANSPOSE',
                   'VLOOKUP']

math_formulas = ['ABS', 'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC',
                 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'BASE', 'CEILING',
                 'CEILING.MATH', 'CEILING.PRECISE', 'COMBIN', 'COMBINA', 'COS',
                 'COSH', 'COT', 'COTH', 'CSC', 'CSCH', 'DECIMAL', 'DEGREES',
                 'EVEN', 'EXP', 'FACT', 'FACTDOUBLE', 'FLOOR', 'FLOOR.MATH',
                 'FLOOR.PRECISE', 'GCD', 'INT', 'ISO.CEILING', 'LCM', 'LN', 'LOG',
                 'LOG10', 'MDETERM', 'MINVERSE', 'MMULT', 'MOD', 'MROUND',
                 'MULTINOMIAL', 'MUNIT', 'ODD', 'PI', 'POWER', 'PRODUCT', 'QUOTIENT',
                 'RADIANS', 'RAND', 'RANDBETWEEN', 'ROMAN', 'ROUND', 'ROUNDDOWN',
                 'ROUNDUP', 'SEC', 'SECH', 'SERIESSUM', 'SIGN', 'SIN', 'SINH',
                 'SQRT', 'SQRTPI', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT',
                 'SUMSQ', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'TAN', 'TANH', 'TRUNC']

statistical_formulas = ['AVEDEV', 'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS',
                        'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE',
                        'BINOM.INV', 'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV',
                        'CHISQ.INV.RT', 'CHISQ.TEST', 'CONFIDENCE.NORM', 'CONFIDENCE.T',
                        'CORREL', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF',
                        'COUNTIFS', 'COVARIANCE.P', 'COVARIANCE.S', 'DEVSQ',
                        'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT',
                        'F.TEST', 'FISHER', 'FISHERINV', 'FORECAST', 'FORECAST.ETS',
                        'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY',
                        'FORECAST.ETS.STAT', 'FORECAST.LINEAR', 'FREQUENCY',
                        'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN', 'GAMMALN.PRECISE',
                        'GAUSS', 'GEOMEAN', 'GROWTH', 'HARMEAN', 'HYPGEOM.DIST',
                        'INTERCEPT', 'KURT', 'LARGE', 'LINEST', 'LOGEST',
                        'LOGNORM.DIST', 'LOGNORM.INV', 'MAX', 'MAXA', 'MEDIAN',
                        'MIN', 'MINA', 'MODE.MULT', 'MODE.SNGL', 'NEGBINOM.DIST',
                        'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV',
                        'PEARSON', 'PERCENTILE.EXC', 'PERCENTILE.INC',
                        'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUT', 'PERMUTATIONA',
                        'PHI', 'POISSON.DIST', 'PROB', 'QUARTILE.EXC', 'QUARTILE.INC',
                        'RANK.AVG', 'RANK.EQ', 'RSQ', 'SKEW', 'SKEW.P', 'SLOPE',
                        'SMALL', 'STANDARDIZE', 'STDEV.P', 'STDEV.S', 'STDEVA',
                        'STDEVPA', 'STEYX', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT',
                        'T.INV', 'T.INV.2T', 'T.TEST', 'TREND', 'TRIMMEAN', 'VAR.P',
                        'VAR.S', 'VARA', 'VARPA', 'WEIBULL.DIST', 'Z.TEST']

text_formulas = ['ASC', 'BAHTTEXT', 'CHAR', 'CLEAN', 'CODE', 'CONCATENATE',
                 'DBCS', 'DOLLAR', 'EXACT', 'FIND', 'FINDB', 'FIXED', 'LEFT',
                 'LEFTB', 'LEN', 'LENB', 'LOWER', 'MID', 'MIDB', 'NUMBERVALUE',
                 'PHONETIC', 'PROPER', 'REPLACE', 'REPLACEB', 'REPT', 'RIGHT',
                 'RIGHTB', 'SEARCH', 'SEARCHB', 'SUBSTITUTE', 'T', 'TEXT',
                 'TRIM', 'UNICHAR', 'UNICODE', 'UPPER', 'VALUE']

user_formulas = ['CALL', 'EUROCONVERT', 'REGISTER.ID', 'SQL.REQUEST']

web_formulas = ['ENCODEURL', 'FILTERXML', 'WEBSERVICE']

csv_columns_list = ['compatibility_formulas', 'cube_formulas', 'database_formulas',
                    'datetime_formulas', 'engineering_formulas', 'financial_formulas',
                    'information_formulas', 'logical_formulas', 'lookup_formulas',
                    'math_formulas', 'statistical_formulas', 'text_formulas',
                    'user_formulas', 'web_formulas']


def get_formula_list():
    formula_list = []
    formula_list.extend(compatibility_formulas)
    formula_list.extend(cube_formulas)
    formula_list.extend(database_formulas)
    formula_list.extend(datetime_formulas)
    formula_list.extend(engineering_formulas)
    formula_list.extend(financial_formulas)
    formula_list.extend(information_formulas)
    formula_list.extend(logical_formulas)
    formula_list.extend(lookup_formulas)
    formula_list.extend(math_formulas)
    formula_list.extend(statistical_formulas)
    formula_list.extend(text_formulas)
    formula_list.extend(user_formulas)
    formula_list.extend(web_formulas)

    return formula_list
