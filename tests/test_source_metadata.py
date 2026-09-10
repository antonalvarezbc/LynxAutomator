import tempfile
import unittest
import zipfile
from pathlib import Path
from lynx_bulk import wi_rows, source_value, default_fields, build_table
from lynx_tasks import TaskContext


class SourceMetadataTests(unittest.TestCase):
    def test_wi_project_foreign_ids_and_partitioned_csvs_join_without_dp_fields(self):
        with tempfile.TemporaryDirectory() as directory:
            archive = Path(directory) / 'wi.zip'
            with zipfile.ZipFile(archive, 'w') as out:
                out.writestr('images_p.csv', 'project_id,deployment_id,image_id,camera_id,location,timestamp,genus,species,latitude,longitude\np,d,i,c,gs://bucket/photo.jpg,2025-01-01,Lynx,pardinus,99,99\n')
                out.writestr('deployments.csv', 'project_id,deployment_id,camera_id,placename,latitude,longitude,timestamp\np,d,c,North,37,-6,2024-01-01\n')
                out.writestr('projects_p.csv', 'project_id,project_admin_id,organization_id,project_name\np,admin,org,Correct project\n')
                out.writestr('projects_other.csv', 'project_id,project_admin_id,organization_id,project_name\nother,admin,org,Wrong project\n')
                out.writestr('organizations.csv', 'organization_id,organization_name\norg,Organization\n')
                out.writestr('cameras.csv', 'camera_id,model_id,camera_name\nc,m,Camera\nother,wrong,Other camera\n')
                out.writestr('models.csv', 'model_id,camera_type\nm,Infrared\nwrong,Wrong\n')
            rows, _ = wi_rows(TaskContext(), archive)
            metadata = rows[0]['metadata']
            self.assertEqual(metadata['projects.project_name'], ['Correct project'])
            self.assertEqual(metadata['organizations.organization_name'], ['Organization'])
            self.assertEqual(metadata['models.camera_type'], ['Infrared'])
            self.assertIn('images.camera_id', metadata)
            self.assertIn('deployments.camera_id', metadata)
            self.assertFalse(any(key.startswith(('media.', 'deployment.', 'observation.', 'package.')) for key in metadata))
            self.assertEqual(source_value(rows[0], 'media.image_id'), 'i')  # old profile alias
            self.assertEqual(source_value(rows[0], 'deployment.placename'), 'North')
            self.assertEqual(source_value(rows[0], 'media.placename'), 'North')
            self.assertEqual(source_value(rows[0], 'media.camera_id_image'), 'c')
            self.assertEqual(source_value(rows[0], 'media.camera_id_deployment'), 'c')
            self.assertEqual(rows[0]['year'], 2025)
            self.assertEqual(rows[0]['latitude'], '37')
            self.assertEqual(rows[0]['longitude'], '-6')
            fields = default_fields() + [dict(name='Encounter.researcherComments', source='projects.project_name', type='text')]
            self.assertEqual(build_table(rows, fields)[0]['Encounter.researcherComments'], 'Correct project')

    def test_dp_gzip_extra_tables_keep_dp_sources(self):
        import gzip
        import json
        from camtrap_factory import make_package
        from lynx_camtrap import read_package, select_media, acquire_media
        from lynx_bulk import camtrap_rows
        with tempfile.TemporaryDirectory() as directory:
            root = make_package(Path(directory) / 'source')
            descriptor = json.loads((root / 'datapackage.json').read_text())
            descriptor['resources'].append({'name': 'cameras', 'path': 'cameras.csv.gz'})
            (root / 'cameras.csv.gz').write_bytes(gzip.compress(b'deploymentID,cameraModel\nd1,Infrared\n'))
            (root / 'datapackage.json').write_text(json.dumps(descriptor))
            task = TaskContext()
            package = read_package(task, root / 'datapackage.json')
            records, _ = select_media(task, package, {'Lynx pardinus'}, True)
            batch = acquire_media(task, package, records, directory, local_only=True)
            rows, _ = camtrap_rows(task, package, records, [batch.output_directory], False)
            self.assertTrue(rows)
            for row in rows:
                self.assertEqual(row['metadata']['cameras.cameraModel'], ['Infrared'])
                self.assertIn('deployment.deploymentID', row['metadata'])
                self.assertIn('media.mediaID', row['metadata'])
                self.assertFalse(any(key.startswith(('images.', 'deployments.', 'projects.')) for key in row['metadata']))
