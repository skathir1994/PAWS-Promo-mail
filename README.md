# PAWS-Promo-mail
PAWS Suspicious Deal Price Notification
Problem: As part of our PAWS error prevention program workflow, used to send notification emails to deal owners on suspicious deal prices. Currently, emails are sent with an impacted set of ASINs attached as an excel sheet containing additional information about the promos. Team recently received a recommendation from an AU stakeholder to present the impacted ASINs and their corresponding details as a table in the email body for quick evaluation and action.

Proposed Solution: However, it is not entirely viable to include all the ASINs in the email body because our notification emails are not limited to a particular set of ASINs. Consequently, I decided to connect the mail body to the message with a maximum of 10 ASINs arranged in order of CPPU impact. Post stakeholder approval, I developed Python code that performs the following logic.

The input file location (PAWS error file) was provided, and the PAWS error was read as a CSV file.

    IF:  The PAWS error exceeds ten. The excel sheet is provided and shared to the deal owners. 
Also displayed in the email body are the top ten forward_looking_cp contributions data.

    Else: The table is displayed in the email body, along with the attached excel sheet, which is shared with the deal owners. 

Impact:  This automated trigger helped increase the deal owner's acknowledgement rate from 53% to 78%, as well as the ease of validation across WW MKPLs.
