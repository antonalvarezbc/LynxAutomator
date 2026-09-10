"""Tiny synthetic WI export, created only in temporary test directories."""
import zipfile
from pathlib import Path
import pandas as pd


def make_wi_zip(folder):
    path = Path(folder) / 'wi.zip'
    images = pd.DataFrame([
        dict(project_id='p', deployment_id='d', image_id='1', location='gs://bucket/lynx.jpg',
             timestamp='2024-01-02T03:04:05Z', genus='Lynx', species='pardinus'),
        dict(project_id='p', deployment_id='d', image_id='2', location='gs://bucket/fox.png',
             timestamp='2024-01-02T03:04:06Z', genus='Vulpes', species='vulpes')])
    deployments = pd.DataFrame([dict(project_id='p', deployment_id='d', placename='Test site')])
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('export/images_p.csv', images.to_csv(index=False))
        archive.writestr('export/deployments.csv', deployments.to_csv(index=False))
        archive.writestr('export/projects.csv', 'project_id,project_name\np,Synthetic project\n')
    return path
