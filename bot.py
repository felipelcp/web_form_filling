"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""


# Import for the Web Bot
from botcity.web import WebBot, Browser, By

import pandas

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    # bot.browser = Browser.FIREFOX

    # Uncomment to set the WebDriver path
    bot.driver_path = "C:\chromedriver-win64\chromedriver.exe"

    # Open Excel file
    data = pandas.read_excel(r"C:\EmployeesFeedback.xlsx")

    # Opens the Form example website.
    bot.browse("https://docs.google.com/forms/d/e/1FAIpQLSf2EwKKGsW7jWBxCNNiJoVEn2vnv9-lcygkBsMuCtsGlKfiEA/viewform")
    # Maximize the window
    bot.maximize_window()
        
    # Wait 3 seconds
    bot.wait(3000)

    # For each row in the Excel file
    for index, row in data.iterrows():

        # Write the name of the Employee
        employee_name_field = bot.find_element("//div[contains(@data-params,'Employee name')]//input[@type='text']", By.XPATH)
        employee_name_field.send_keys(row['Employee Name'])
        print(row['Employee Name'])

        # Wait 1 seconds
        bot.wait(1000)

        # Write the years of service
        years_of_service_field = bot.find_element("//div[contains(@data-params,'Years of Service')]//input[@type='text']", By.XPATH)
        years_of_service_field.send_keys(row['Years of Service'])
        print(row['Years of Service'])

        # Wait 1 seconds
        bot.wait(1000)

        # Select the Department
        department_field = bot.find_element("//div[contains(@data-params, 'Department')]//div[contains(@role,'listbox')]", By.XPATH)
        department_field.click()
        bot.wait(1000)
        department_field_option = bot.find_element("//div[@role='option' and @data-value='"+row['Department']+"']", By.XPATH)
        department_field_option.click()
        print(row['Department'])

        # Wait 1 seconds
        bot.wait(1000)

        # Select the employee satisfaction
        employee_satisfaction_field = bot.find_element("//div[contains(@data-params, 'Employee satisfaction rating')]//span[text()='"+row['Satisfaction Rating']+"']", By.XPATH)
        employee_satisfaction_field.click()
        print(row['Satisfaction Rating'])

        # Wait 1 seconds
        bot.wait(1000)

        # Click in the "Enviar" button
        submit_btn = bot.find_element("//div[@role='button']//span[text()='Enviar']", by="xpath")
        submit_btn.click()

        # Wait 1 seconds
        bot.wait(1000)

        # Click in the "Enivar outra resposta" link
        submit_another_response_btn = bot.find_element("//a[text()='Enviar outra resposta']", by="xpath")
        submit_another_response_btn.click()

        # Wait 1 seconds
        bot.wait(1000)

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finished OK."
    )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
