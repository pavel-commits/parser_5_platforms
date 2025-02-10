import sqlite3

TABLES = {
    "zacks_screener": [('Ticker', '''TEXT PRIMARY KEY
                                UNIQUE
                                NOT NULL'''),
                       ('Market Cap (mil)', 'REAL'),
                       ('Last EPS Report Date (yyyymmdd)', 'INTEGER'),
                       ('Next EPS Report Date  (yyyymmdd)', 'INTEGER'),
                       ('% Change Q1 Est. (4 weeks)', 'REAL'),
                       ('% Change Q2 Est. (4 weeks)', 'REAL'),
                       ('% Change F1 Est. (4 weeks)', 'REAL'),
                       ('% Change F2 Est. (4 weeks)', 'REAL'),
                       ('Q1 Consensus Est. ', 'REAL'),
                       ('Q2 Consensus Est. (next fiscal Qtr)', 'REAL'),
                       ('F1 Consensus Est.', 'REAL'),
                       ('F2 Consensus Est.', 'REAL'),
                       ('# of Analysts in Q1 Consensus', 'INTEGER'),
                       ('F(1) Consensus Sales Est. ($mil)', 'REAL'),
                       ('Q(1) Consensus Sales Est. ($mil)', 'REAL'),
                       ('# of Analysts in F1 Consensus', 'INTEGER'),
                       ('# of Analysts in F2 Consensus', 'INTEGER'),
                       ('St. Dev. Q1 / Q1 Consensus', 'TEXT'),
                       ('St. Dev. F1 / F1 Consensus', 'TEXT')],
    "zacks_detailed": [('Ticker', '''TEXT    PRIMARY KEY
                                        UNIQUE
                                        NOT NULL'''),
                       ('Sales_Current_high', 'REAL'),
                       ('Sales_Current_low', 'REAL'),
                       ('Sales_Next_high', 'REAL'),
                       ('Sales_Next_low', 'REAL'),
                       ('Earnings_Current_high', 'REAL'),
                       ('Earnings_Current_low', 'REAL'),
                       ('Earnings_Next_high', 'REAL'),
                       ('Earnings_Next_low', 'REAL')],
    "zacks_calendar": [('Ticker', 'TEXT     NOT NULL'),
                       ('Date', 'DATE'),
                       ('Surprise', 'TEXT'),
                       ('Time', 'TEXT')],
    "zacks_consensus": [('Ticker', 'TEXT NOT NULL'),
                        ('date', 'DATE'),
                        ('value', 'REAL')],
    "saver_ms": [('Ticker', 'TEXT NOT NULL'),
                 ('2024_Book Value Per Share', 'REAL'),
                 ('2024_Net sales', "REAL"),
                 ('2024_EBITDA', "REAL"),
                 ('2024_EBIT', "REAL"),
                 ('2024_Net income', "REAL"),
                 ('2024_EPS', "REAL"),
                 ('2025_Book Value Per Share', 'REAL'),
                 ('2025_Net sales', "REAL"),
                 ('2025_EBITDA', "REAL"),
                 ('2025_EBIT', "REAL"),
                 ('2025_Net income', "REAL"),
                 ('2025_EPS', "REAL"),
                 ('2024_Q1_EPS', "REAL"),
                 ('2024_Q2_EPS', "REAL"),
                 ('oper_date', "DATE")],
    "saver_fmp": [('Ticker', 'TEXT NOT NULL'),
                  ('AvgPriceTarget', 'REAL'),
                  ('Rec 2m', 'REAL'),
                  ('oper_date', "DATE")],
    "saver_yh": [('Ticker', 'TEXT NOT NULL'),
                 ('shares_short', 'REAL'),
                 ('epsEstimate', 'REAL'),
                 ('oper_date', "DATE")],
    "marketscreener": [('Ticker', '''TEXT    PRIMARY KEY
                                             UNIQUE
                                             NOT NULL'''),
                       ('Fiscal period', 'TEXT'),
                       ('Currency', 'TEXT'),
                       ('Size', 'TEXT'),
                       ('2024_Net Debt', 'REAL'),
                       ('2024_Net Cash position', 'REAL'),
                       ('2024_Assets', 'REAL'),
                       ('2024_Cash Flow per Share', 'REAL'),
                       ('2024_Capex', 'REAL'),
                       ('2024_Earnings before Tax (EBT)', 'REAL'),
                       ('2024_Free Cash Flow', 'REAL'),
                       ('2024_Dividend per Share', 'REAL'),
                       ('2025_Net Debt', 'REAL'),
                       ('2025_Net Cash position', 'REAL'),
                       ('2025_Assets', 'REAL'),
                       ('2025_Cash Flow per Share', 'REAL'),
                       ('2025_Capex', 'REAL'),
                       ('2025_Earnings before Tax (EBT)', 'REAL'),
                       ('2025_Free Cash Flow', 'REAL'),
                       ('2025_Dividend per Share', 'REAL'),
                       ('2024_Q1_Net sales', 'REAL'),
                       ('2024_Q1_EBITDA', 'REAL'),
                       ('2024_Q1_EBIT', 'REAL'),
                       ('2024_Q1_Earnings before Tax (EBT)', 'REAL'),
                       ('2024_Q1_Net income', 'REAL'),
                       ('2024_Q1_Dividend per Share', 'REAL'),
                       ('2024_Q1_Announcement Date', 'REAL'),
                       ('2024_Q2_Net sales', 'REAL'),
                       ('2024_Q2_EBITDA', 'REAL'),
                       ('2024_Q2_EBIT', 'REAL'),
                       ('2024_Q2_Earnings before Tax (EBT)', 'REAL'),
                       ('2024_Q2_Net income', 'REAL'),
                       ('2024_Q2_Dividend per Share', 'REAL'),
                       ('2024_Q2_Announcement Date', 'REAL'),
                       ('2024_Q3_Net sales', 'REAL'),
                       ('2024_Q3_EBITDA', 'REAL'),
                       ('2024_Q3_EBIT', 'REAL'),
                       ('2024_Q3_Earnings before Tax (EBT)', 'REAL'),
                       ('2024_Q3_Net income', 'REAL'),
                       ('2024_Q3_EPS', 'REAL'),
                       ('2024_Q3_Dividend per Share', 'REAL'),
                       ('2024_Q3_Announcement Date', 'REAL'),
                       ('2024_Q4_Net sales', 'REAL'),
                       ('2024_Q4_EBITDA', 'REAL'),
                       ('2024_Q4_EBIT', 'REAL'),
                       ('2024_Q4_Earnings before Tax (EBT)', 'REAL'),
                       ('2024_Q4_Net income', 'REAL'),
                       ('2024_Q4_EPS', 'REAL'),
                       ('2024_Q4_Dividend per Share', 'REAL'),
                       ('2024_Q4_Announcement Date', 'REAL')],
    "yahoo": [('Ticker', '''TEXT    PRIMARY KEY
                                        UNIQUE
                                        NOT NULL'''),
              ('1yTargetEst', 'REAL'),
              ('fullTimeEmployees', 'REAL'),
              ('averageDailyVolume3Month', 'REAL'),
              ('shortPercentOfFloat', 'REAL'),
              ('sharesShortPriorMonth', 'REAL'),
              ('trailingAnnualDividendRate', 'REAL'),
              ('epsEstimate', 'REAL'),
              ('epsActual', 'REAL'),
              ('currentEstimate', 'REAL'),
              ('7daysAgo', 'REAL'),
              ('30daysAgo', 'REAL'),
              ('currentYear', 'REAL'),
              ('nextYear', 'REAL'),
              ('next5Years', 'REAL'),
              ('past5Years', 'REAL'),
              ('corporateGovernance', 'REAL'),
              ('totalEsg', 'REAL'),
              ('environmentScore', 'REAL'),
              ('socialScore', 'REAL'),
              ('governanceScore', 'REAL'),
              ('highestControversy', 'REAL')],
    "yahoo_an": [('Ticker', '''TEXT    PRIMARY KEY
                                       UNIQUE
                                       NOT NULL'''),
                 ('numberOfAnalysts_Qtr_cur', 'REAL'),
                 ('numberOfAnalysts_Qtr_next', 'REAL'),
                 ('numberOfAnalysts_Year_cur', 'REAL'),
                 ('numberOfAnalysts_Year_next', 'REAL'),
                 ('earnings_avg_Qtr_cur', 'REAL'),
                 ('earnings_avg_Qtr_next', 'REAL'),
                 ('earnings_avg_Year_cur', 'REAL'),
                 ('earnings_avg_Year_next', 'REAL'),
                 ('low_Qtr_cur', 'REAL'),
                 ('low_Qtr_next', 'REAL'),
                 ('low_Year_cur', 'REAL'),
                 ('low_Year_next', 'REAL'),
                 ('high_Qtr_cur', 'REAL'),
                 ('high_Qtr_next', 'REAL'),
                 ('high_Year_cur', 'REAL'),
                 ('high_Year_next', 'REAL'),
                 ('revenue_avg_Qtr_cur', 'REAL'),
                 ('revenue_avg_Qtr_next', 'REAL'),
                 ('revenue_avg_Year_cur', 'REAL'),
                 ('revenue_avg_Year_next', 'REAL'),
                 ('salesGrowth_Qtr_cur', 'REAL'),
                 ('salesGrowth_Qtr_next', 'REAL'),
                 ('salesGrowth_Year_cur', 'REAL'),
                 ('salesGrowth_Year_next', 'REAL')]
}


def zacks_screener(cur):
    f3 = TABLES.get("zacks_screener")

    r = f"""
            CREATE TABLE zacks_screener (
            {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
        );
        """
    print(r)
    cur.execute(r)


def zacks_detailed(cur):
    f3 = TABLES.get("zacks_detailed")

    r = f"""
            CREATE TABLE zacks_detailed (
            {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
        );
        """
    # print(r)
    cur.execute(r)


def zacks_calendar(cur):
    f3 = TABLES.get("zacks_calendar")

    r = f"""
                CREATE TABLE zacks_calendar (
                {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
            );
            """
    # print(r)
    cur.execute(r)

    r = f"""
        INSERT INTO zacks_calendar VALUES
            (?, ?, ?, ?);
    """
    # print(r)
    # for i in range(1000000):
    #     cur.execute(r, ('AA' + str(i), datetime.date.today(), 5.63 + i, "After close"))


def zacks_consensus(cur):
    f3 = TABLES.get("zacks_consensus")

    r = f"""CREATE TABLE zacks_consensus (
                {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
            );
            """
    # print(r)
    cur.execute(r)

    # r = f"""
    #     INSERT INTO zacks_consensus VALUES
    #         (?, ?, ?);
    # """
    # cur.execute(r, ('AA', datetime.date.today(), 5.63))


def saver_ms(cur):
    f3 = TABLES.get("saver_ms")
    # fmp
    #             ('AvgPriceTarget',  'REAL'),

    r = f"""CREATE TABLE saver_ms (
                {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
            );
            """
    # print(r)
    cur.execute(r)


def saver_fmp(cur):
    f3 = TABLES.get("saver_fmp")

    r = f"""CREATE TABLE saver_fmp (
                {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
            );
            """
    # print(r)
    cur.execute(r)


def saver_yh(cur):
    f3 = TABLES.get("saver_yh")

    r = f"""CREATE TABLE saver_yh (
                {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
            );
            """
    # print(r)
    cur.execute(r)


def marketscreener(cur):
    f3 = TABLES.get("marketscreener")

    r = f"""CREATE TABLE marketscreener (
                    {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
                );
                """
    # print(r)
    cur.execute(r)


def yahoo(cur):
    f3 = TABLES.get("yahoo")

    r = f"""CREATE TABLE yahoo (
                    {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
                );
                """
    # print(r)
    cur.execute(r)


def yahoo_an(cur):
    f3 = TABLES.get("yahoo_an")

    r = f"""CREATE TABLE yahoo_an (
                    {', '.join([f"[{t[0]}] {t[1]}" for t in f3])}
                );
                """
    # print(r)
    cur.execute(r)


def main():
    con = sqlite3.connect("all_analyst.db")
    cur = con.cursor()

    # zacks_screener(cur)
    # zacks_detailed(cur)
    # zacks_calendar(cur)
    # zacks_consensus(cur)
    # marketscreener(cur)
    # saver_ms(cur)
    # saver_fmp(cur)
    # saver_yh(cur)
    # marketscreener(cur)
    # yahoo(cur)
    # yahoo_an(cur)

    con.close()
