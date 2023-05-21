This Lambda Python code will generate weekly costing, cost anomalies, Reseved Instance recommendations across set of AWS accounts and sends an email to AWS account owners to have a quick review and summery of the past week usage in costing perspective. Also this Lambda will send Excel files as attachments with all the costing details.

You need to setup following variables in the lambda:
    - ASSUME_ROLE [The name of the IAM role that has permission to AWS get_cost_and_usage, budgets, reservation_purchase_recommendation, and anomalies]
    - SES_REGION [The region of the email recipents verified in AWS SES]
    - SEND_FROM [The email addres this report coming from]