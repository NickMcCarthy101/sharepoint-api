import json
import requests 
import sys
import os
import re 
from urllib.parse import urlparse, parse_qs
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.utilities.request_options import RequestOptions

USERNAME = ''
PASSWORD = ''
BASE_URL = ''
ctx = AuthenticationContext(url=BASE_URL)

SP_RE = re.compile(".+:\/\/(.+).sharepoint.com(.+)Shared%20Documents\/Forms\/AllItems.aspx")
def get_endpoint_from_url(url):
    match = re.match(SP_RE, url)
    if not match:
        return ''
    sp_site = match.group(1)
    sp_group = match.group(2)
    query = parse_qs(urlparse(url).query)
    path = query.get('id', ['Shared Documents'])[0]
    return f"https://{sp_site}.sharepoint.com{sp_group}_api/web/GetFolderByServerRelativeUrl('{path}')", path

def retrieve_file_list(url):
    url = url + '/Files'
    if ctx.acquire_token_for_user(username=USERNAME, password=PASSWORD):
        options = RequestOptions(BASE_URL)
        context = ClientContext(BASE_URL, ctx)
        context.request_form_digest()
        options.set_header('Accept', 'application/json; odata=verbose')
        options.method = 'GET'
        context.authenticate_request(options)
        data = requests.get(url=url, headers=options.headers, auth=options.auth)
        if data.status_code == 200:
            file_list = []
            datam = json.loads(data.text)
            for f in range(len(datam['d']['results'])):
                uri = datam['d']['results'][f]['__metadata']['uri']
                file_list.append(uri)
            return file_list
        else:
            print(data.content)

def download_file(uri):
    """Give the URI (returned from retrieve file query) 
    and the path you want to write to."""
    uriR = uri + '/$value'
    if ctx.acquire_token_for_user(username=USERNAME, password=PASSWORD):
        options = RequestOptions(BASE_URL)
        context = ClientContext(BASE_URL, ctx)
        context.request_form_digest()
        options.method = 'GET'
        context.authenticate_request(options)
        byte_doc = requests.get(url=uriR, headers=options.headers, auth=options.auth)
        hack_file_name = uri.split("/")
        hack_file = hack_file_name[len(hack_file_name)-1]
        file_name = hack_file[0:len(hack_file)-2]
        return file_name, byte_doc.content
    else:
        print('Incorrect login credentials')

def upload_file_from_local(local_file_path, sharepoint_folder_path):
    """Upload a file to a folder give file name and folder path
       The sharepoint_folder_path is any subdirectory in Shared Documents/"""
    if ctx.acquire_token_for_user(username=USERNAME, password=PASSWORD):
        s_folder = sharepoint_folder_path
        fname = os.path.basename(os.path.normpath(local_file_path))
        files_url ="{0}/_api/web/GetFolderByServerRelativeUrl('{1}')/Files/add(url='{2}', overwrite=true)"
        full_url = files_url.format(BASE_URL, s_folder, fname)
        options = RequestOptions(BASE_URL)
        context = ClientContext(BASE_URL, ctx)
        context.request_form_digest()
        options.set_header('Accept', 'application/json; odata=verbose')
        options.set_header('Content-Type', 'application/octet-stream')
        options.set_header('Content-Length', str(os.path.getsize(local_file_path)))
        options.set_header('X-RequestDigest', context.contextWebInformation.form_digest_value)
        options.method = 'POST'

    with open(local_file_path, 'rb') as outfile:
        context.authenticate_request(options)
        data = requests.post(url=full_url, data=outfile, headers=options.headers, auth=options.auth)
        if data.status_code == 200:
            print(data.ok)
        else:
            print(data.content)

def create_sharepoint_folder(new_folder_name):
    """THIS API CAN ONLY HANDLE 1 FOLDER CREATION AT A TIME"""
    if ctx.acquire_token_for_user(username=USERNAME, password=PASSWORD):
        create_folder_url ="{0}/_api/web/GetFolderByServerRelativeUrl('Shared Documents')/folders".format(BASE_URL)
        options = RequestOptions(BASE_URL)
        context = ClientContext(BASE_URL, ctx)
        context.request_form_digest()
        options.set_header('Accept', 'application/json; odata=verbose')
        options.set_header('Content-Type', 'application/json;odata=verbose')
        options.set_header('X-RequestDigest', context.contextWebInformation.form_digest_value)
        options.method = 'POST'
        context.authenticate_request(options)
        body = {}
        body['__metadata'] = {'type': 'SP.Folder'}
        body['ServerRelativeUrl'] = new_folder_name
        string_body = json.dumps(body)
        data = requests.post(url=create_folder_url, data=string_body, headers=options.headers, auth=options.auth)
        return data.ok

