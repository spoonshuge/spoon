import logging
import azure.functions as func

from azure.identity import DefaultAzureCredential
from msgraph.core import GraphClient

# create a single, reusable Graph client
credential = DefaultAzureCredential()
graph_client = GraphClient(credential=credential, scopes=["https://graph.microsoft.com/.default"])

# constants for your file location
# the host name from your URL:
HOSTNAME = "danielrowlandgmail-my.sharepoint.com"
# the path under /personal/... to your file
PERSONAL_PATH = "/personal/daniel_connecteduniverse_ai/Documents"
FILENAME = "sLcGiveaway.xlsx"
TABLE_NAME = "Table1"

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="http_trigger_giveaway")
def http_trigger_giveaway(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Python HTTP trigger processed a request.")

    # pull the 'name'
    name = req.params.get("name")
    if not name:
        try:
            body = req.get_json()
        except ValueError:
            body = {}
        name = body.get("name")

    if not name:
        return func.HttpResponse(
            "Pass a name in the query string or in the request body.",
            status_code=400
        )

    # 1️⃣ locate your OneDrive site
    site = graph_client.get(
        f"/sites/{HOSTNAME}:/personal/daniel_connecteduniverse_ai"
    ).json()
    site_id = site["id"]

    # 2️⃣ get the drive (OneDrive) under that site
    drive = graph_client.get(f"/sites/{site_id}/drive").json()
    drive_id = drive["id"]

    # 3️⃣ locate the file by path
    item = graph_client.get(
        f"/drives/{drive_id}/root:{PERSONAL_PATH}/{FILENAME}"
    ).json()
    item_id = item["id"]

    # 4️⃣ append a new row to Table1
    payload = {
        "values": [[ name ]]
    }
    graph_client.post(
        f"/drives/{drive_id}/items/{item_id}/workbook/tables/{TABLE_NAME}/rows/add",
        json=payload
    )

    return func.HttpResponse(f"Added '{name}' to your giveaway sheet! 👍")