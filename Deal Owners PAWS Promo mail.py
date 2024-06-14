# Importing the necessary packages
import pandas as pd
import win32com.client

df = pd.read_csv(r'C:\Users\skathir\Desktop\PR File\2024\APRIL\test2.csv')
newdf = df["deal_owner"].drop_duplicates()
ls = newdf.tolist()

# Using Zip method
mang = zip(df.deal_owner, df.manager_id)

# Converting it into list
mang = list(set(mang))
# print(mang)

l = len(df)
# print(l)
#
# Creating new DF.
#
A_table = pd.DataFrame(
    data=df,
    columns=[
        "Week",
        "marketplace",
        "deal_id",
        "paws_id",
        "deal_start_date",
        "deal_owner",
        "pf_desc",
        "gl_desc",
        "asin",
        "item_name",
        "deal_price",
        "current_price_with_tax",
        "current_price_no_tax",
        "lowest_price_l30d_with_tax",
        "lowest_price_ytd_with_tax",
    ],
)

#
# Read the relevant text file for the outlook body.
#

with open(
    r'C:\Users\skathir\PycharmProjects\Deal Owners PAWS Promotion\table style.txt'
) as file:
    table_style = file.read()

with open(
    r'C:\Users\skathir\PycharmProjects\Deal Owners PAWS Promotion\before table html.txt'
) as body_file:
    body_html = body_file.read()

with open(
    r'C:\Users\skathir\PycharmProjects\Deal Owners PAWS Promotion\after_table_html.txt',
    errors="ignore",
) as body_file:
    last_body_html = body_file.read()

html = A_table.to_html(index=False)
# write html to file
text_file = open("index.html", "w")
text_file.write(html)
text_file.close()
# print()

# Converting it into dictionary
mang = dict(mang)
# print(mang)
for m in ls:
    maildf = df[
        [
            "Week",
            "marketplace",
            "deal_id",
            "paws_id",
            "deal_start_date",
            "deal_owner",
            "pf_desc",
            "gl_desc",
            "asin",
            "item_name",
            "deal_price",
            "current_price_with_tax",
            "current_price_no_tax",
            "lowest_price_l30d_with_tax",
            "lowest_price_ytd_with_tax",
            "forward_looking_cp",
            "fully_loaded_cost_amount",
            "manager_id",
        ]
    ][df["deal_owner"] == m].round(2)
    wk = maildf["Week"].drop_duplicates()
    s = [str(i) for i in wk]
    week = " & ".join(s)
    mk = maildf["marketplace"].drop_duplicates()
    mkpl = " & ".join(mk)
    g = maildf["gl_desc"].drop_duplicates()

    # Sorting the values based on FL_CP.
    f_cp = maildf.sort_values(by=["forward_looking_cp"], ascending=False).head(10)
    html0 = f_cp.to_html(index=False)
    # write html to file
    text_file = open("index.html0", "w")
    text_file.write(html0)
    text_file.close()

    if len(g) > 2:
        p = maildf["pf_desc"].drop_duplicates()
        gl = " & ".join(p)
        gl = gl + ' PL'
    else:
        gl = " & ".join(g)
        gl = gl + ' GL'
    #
    # More then 10 listings.
    #
    if l > 10:
        fname = (
            'Week '
            + week
            + '_'
            + mkpl
            + ' MKPL_'
            + gl
            + '_Suspicious PAWS Promotion Price'
        )
        maildf.to_csv(
            'C:/Users/skathir/Desktop/PR File/2024/APRIL/' + fname + '.csv', index=False
        )
        ol = win32com.client.Dispatch("outlook.application")
        olmailitem = 0x0  # size of the new email
        newmail = ol.CreateItem(olmailitem)
        newmail.Subject = (
            'Action Needed: Suspicious PAWS Promotion Price_Week '
            + week
            + '_'
            + mkpl
            + ' MKPL_'
            + gl
            + ''
        )
        #   newmail.To=m+'@amazon.com'
        newmail.To = 'skathir@amazon.com'
        #   newmail.CC='ycchoudh@amazon.com;cmt-rp-fnap@amazon.com;'+mang[m]+'@amazon.com'
        newmail.HTMLBody = body_html + '<br/> PAWS Promo <br/>' + html0 + last_body_html
        attach = 'C:/Users/skathir/Desktop/PR File/2024/APRIL/' + fname + '.csv'
        newmail.Attachments.Add(attach)
        newmail.Send()
    #
    # Less then 10 listings.
    #
    else:
        fname = (
            'Week '
            + week
            + '_'
            + mkpl
            + ' MKPL_'
            + gl
            + '_Suspicious PAWS Promotion Price'
        )
        maildf.to_csv(
            'C:/Users/skathir/Desktop/PR File/2024/APRIL/' + fname + '.csv', index=False
        )
        ol = win32com.client.Dispatch("outlook.application")
        olmailitem = 0x0  # size of the new email
        newmail = ol.CreateItem(olmailitem)
        newmail.Subject = (
            'Action Needed: Suspicious PAWS Promotion Price_Week '
            + week
            + '_'
            + mkpl
            + ' MKPL_'
            + gl
            + ''
        )
        newmail.HTMLBody = body_html + '<br/> PAWS Promo <br/>' + html + last_body_html
        # newmail.To=m+'@amazon.com'
        newmail.To = 'skathir@amazon.com'
        # newmail.CC='ycchoudh@amazon.com;cmt-rp-fnap@amazon.com;'+mang[m]+'@amazon.com'
        attach = 'C:/Users/skathir/Desktop/PR File/2024/APRIL/' + fname + '.csv'
        newmail.Attachments.Add(attach)
        newmail.Send()

print("Succesfully Excecuted")
