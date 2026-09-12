"""Regression test: referenced submodels and properties must survive AASX export."""
import tempfile
import unittest
from pathlib import Path
from basyx.aas import model
from basyx.aas.adapter.aasx import AASXWriter, AASXReader, DictSupplementaryFileContainer

class SerializationTest(unittest.TestCase):
    def test_referenced_submodel_and_value_survive(self):
        sm=model.Submodel(id_='urn:test:sm',id_short='TestSM',submodel_element=[model.Property(id_short='Example',value_type=str,value='expected')])
        aas=model.AssetAdministrationShell(id_='urn:test:aas',id_short='TestAAS',
             asset_information=model.AssetInformation(asset_kind=model.AssetKind.INSTANCE,global_asset_id='urn:test:asset'),
             submodel={model.ModelReference.from_referable(sm)})
        store=model.DictIdentifiableStore([aas,sm])
        with tempfile.TemporaryDirectory(prefix='aas_serialization_') as tmp:
            path=Path(tmp)/'test.aasx'
            with AASXWriter(path,failsafe=False) as writer:
                writer.write_aas(aas_ids=aas.id,object_store=store,file_store=DictSupplementaryFileContainer(),write_json=True)
            actual=model.DictIdentifiableStore()
            with AASXReader(path,failsafe=False) as reader:reader.read_into(actual,DictSupplementaryFileContainer())
            self.assertEqual(len(actual),2)
            self.assertEqual(actual.get_item(sm.id).submodel_element.get_object_by_attribute('id_short','Example').value,'expected')

if __name__=='__main__':unittest.main()
