import uuid
import json
import requests

class Todoist:
    """ Todoist API wrapper """

    def __init__(self, verify_ssl):
        self.url = ""
        self.token = ""
        self.base_headers = {}
        self.auth_header = {}
        self.url_tasks = ""
        self.url_projects = ""
        self.verify_ssl = verify_ssl

    def _get_tasks_url(self):
        return self.url_tasks

    def _get_projects_url(self):
        return self.url_projects

    def connect(self, url, token):
        self.url = url if url.endswith('/') else url + "/"
        self.token = token
        self.auth_header = {"Authorization": f"Bearer {self.token}"}

        self.url_tasks = self.url + "tasks"
        self.url_projects = self.url + "projects"


    def get_all_projects(self):
        url = self._get_projects_url()
        response = requests.get(url, headers=self.auth_header, verify=self.verify_ssl)

        if response.status_code != 200:
            response.raise_for_status()

        return response.json()


    def delete_task(self, task_id):
        url = self._get_tasks_url() + "/" + str(task_id)
        response = requests.delete(url, headers=self.auth_header, verify=self.verify_ssl)

        if response.status_code != 204:
            response.raise_for_status()


    def delete_tasks(self, filter_project_id, filter_label_id = None):

        task_list = self.get_active_tasks(filter_project_id, filter_label_id)
        for task in task_list:
            #print("Task ID: " + str(task["id"]))
            self.delete_task(task["id"])



    def get_active_tasks(self, filter_project_id, filter_label_id):
        params = {}

        if filter_project_id is not None:
            params["project_id"] = filter_project_id
        if filter_label_id is not None:
            params["label_id"] = filter_label_id

        url = self._get_tasks_url()
        response = requests.get(url, params, headers=self.auth_header, verify=self.verify_ssl)

        if response.status_code != 200:
            response.raise_for_status()

        return response.json()


    def add_new_task(self, project_id, content, datetime_string, label_id):
        """ date_string - date in YYYY-MM-DD format """
        url = self._get_tasks_url()

        params = {"project_id": project_id, "content": content}

        if datetime_string is not None:
            params["due_datetime"] = datetime_string

        if label_id is not None:
            labels_array = [label_id]
            params["label_ids"] = labels_array

        json_data = json.dumps(params)

        hdrs = {"Content-Type": "application/json",
                "X-Request-Id": str(uuid.uuid4()),
                "Authorization": f"Bearer {self.token}" }

        #print(f"Url: {url}")
        #print(f"Hdr: {hdrs}")
        #print(f"Json: {json_data}")

        response = requests.post(url, data = json_data, headers=hdrs, verify=self.verify_ssl)

        if response.status_code != 200:
            print(response.content)
            response.raise_for_status()
