"""
Bluebook Manager — Outsource service (business logic).
"""

from dal import dal
from dal.models import Bluebook, Outsource
from services.log_service import log


def create_outsource(name: str, contact_info: str = "") -> Outsource:
    oid = dal.add_outsource(name, contact_info)
    log("CREATE_OUTSOURCE", f"Name: {name} (id={oid})")
    return dal.get_outsource(oid)


def get_outsource(outsource_id: int) -> Outsource:
    return dal.get_outsource(outsource_id)


def list_outsources() -> list[Outsource]:
    return dal.list_outsources()


def update_outsource(outsource_id: int, name: str, contact_info: str = ""):
    dal.update_outsource(outsource_id, name, contact_info)
    log("UPDATE_OUTSOURCE", f"id={outsource_id}, name={name}")


def delete_outsource(outsource_id: int):
    o = dal.get_outsource(outsource_id)
    if o:
        dal.delete_outsource(outsource_id)
        log("DELETE_OUTSOURCE", f"Name: {o.name} (id={outsource_id})")


def link_bluebook(outsource_id: int, bluebook_id: int):
    dal.link_outsource_bluebook(outsource_id, bluebook_id)
    log("LINK_OUTSOURCE_BLUEBOOK", f"outsource={outsource_id}, bluebook={bluebook_id}")


def unlink_bluebook(outsource_id: int, bluebook_id: int):
    dal.unlink_outsource_bluebook(outsource_id, bluebook_id)
    log("UNLINK_OUTSOURCE_BLUEBOOK", f"outsource={outsource_id}, bluebook={bluebook_id}")


def get_outsources_for_bluebook(bluebook_id: int) -> list[Outsource]:
    return dal.get_outsources_for_bluebook(bluebook_id)
