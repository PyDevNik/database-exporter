from typing import List, Union, Type

from pymongo import MongoClient
from pydantic import BaseModel


class FirstTaskSchema(BaseModel):
    code: str
    documents: List[str]


class SecondTaskSchema(BaseModel):
    code: str
    description: str


class MongoInserter:
    def __init__(self, db_name: str):
        self.__mongo_client = MongoClient()
        self._db = self.__mongo_client.get_database(db_name)

    def launch_insertion(
            self,
            collection_name: str,
            data: List[dict],
            pydantic_schema: Type[Union[FirstTaskSchema, SecondTaskSchema]]
    ) -> None:
        """
        :param collection_name: Collection name in mongo where data will be inserted
        :param data: actual incoming data from xlsx parser
        :param pydantic_schema: Schema for validation
        """
        self.check_data(data, pydantic_schema)
        collection = self._db.get_collection(collection_name)
        collection.insert_many(data)

    @staticmethod
    def check_data(
            data: List[dict],
            pydantic_schema: Type[Union[FirstTaskSchema, SecondTaskSchema]]
    ) -> None:
        """
        Checks whether incoming data is valid

        :param data: incoming data from xlsx parsing
        :param pydantic_schema: Schema to validate objects inside of list
        """
        for datum in data:
            pydantic_schema.parse_obj(datum)

