"""
Bluebook Manager — Customer service (business logic).
"""

from dal import dal
from dal.models import Bluebook, Customer
from services.log_service import log


def create_customer(name: str, contact_info: str = "") -> Customer:
    cid = dal.add_customer(name, contact_info)
    log("CREATE_CUSTOMER", f"Name: {name} (id={cid})")
    return dal.get_customer(cid)


def get_customer(customer_id: int) -> Customer:
    return dal.get_customer(customer_id)


def list_customers() -> list[Customer]:
    return dal.list_customers()


def update_customer(customer_id: int, name: str, contact_info: str = ""):
    dal.update_customer(customer_id, name, contact_info)
    log("UPDATE_CUSTOMER", f"id={customer_id}, name={name}")


def delete_customer(customer_id: int):
    c = dal.get_customer(customer_id)
    if c:
        dal.delete_customer(customer_id)
        log("DELETE_CUSTOMER", f"Name: {c.name} (id={customer_id})")


def link_bluebook(customer_id: int, bluebook_id: int):
    dal.link_customer_bluebook(customer_id, bluebook_id)
    log("LINK_CUSTOMER_BLUEBOOK", f"customer={customer_id}, bluebook={bluebook_id}")


def unlink_bluebook(customer_id: int, bluebook_id: int):
    dal.unlink_customer_bluebook(customer_id, bluebook_id)
    log("UNLINK_CUSTOMER_BLUEBOOK", f"customer={customer_id}, bluebook={bluebook_id}")


def get_bluebooks_for_customer(customer_id: int) -> list[Bluebook]:
    return dal.get_bluebooks_for_customer(customer_id)


def get_customers_for_bluebook(bluebook_id: int) -> list[Customer]:
    return dal.get_customers_for_bluebook(bluebook_id)
