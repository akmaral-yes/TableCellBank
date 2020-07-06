import pandas as pd
import requests
import os
import uuid


def build_url_uuid_df(url_list):
    """
    Given a set of unique urls, build a data frame with the urls and
    corresponding uuids
    """
    url_uuid_df = pd.DataFrame(columns=["url", "uuid"])
    for url in url_list:
        uuid_ = uuid.uuid4()
        add = pd.DataFrame([[url, uuid_]], columns=url_uuid_df.columns)
        url_uuid_df = url_uuid_df.append(add)
    url_uuid_df.reset_index(drop=True, inplace=True)
    return url_uuid_df


def download_files(url_df, folder_to_save):
    """
    Given a data frame with urls and uuids, download the files from the urls
    and save them with corresponding uuid
    """
    for row in url_df.index:
        url = url_df.loc[row, "url"]
        filename = url_df.loc[row, "uuid"]
        filename = str(filename)+".docx"
        file_path = os.path.join(folder_to_save, filename)

        try:
            s = requests.Session()
            r = s.get(url, timeout=10)
            with open(file_path, 'wb') as outfile:
                outfile.write(r.content)
        except BaseException:
            pass


def main():
    """
    Download url.csv from TableBank/TableBank_data/Recogntion_data/Word and
    save it in '../url_docx/url.csv'
    """
    # Load url.csv
    url_path = '../url_docx/url.csv'
    url_df = pd.read_csv(url_path)

    # Get a list of urls
    url_list = url_df['url'].tolist()

    # Get a set of unique urls
    url_list = list(set(url_list))

    # Build a data frame with the urls and corresponding uuids
    url_uuid_df = build_url_uuid_df(url_list)

    # Save the data frame
    url_uuid_path = "../url_docx/url_uuid.csv"
    url_uuid_df.to_csv(url_uuid_path)

    # Create a folder to save downloaded .docx files
    folder_to_save = "../data_docx_structure"
    os.makedirs(folder_to_save, exist_ok=True)

    download_files(url_uuid_df, folder_to_save)


if __name__ == "__main__":
    main()
