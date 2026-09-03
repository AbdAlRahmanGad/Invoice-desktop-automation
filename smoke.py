"""A run of the whole workflow against Fakturama, without the OCR.

Same steps as `main.py` -- an order, its customer, its products, and the
invoice raised from it -- driven from the values extraction reads off
invoice.pdf rather than by reading the document again. That makes it a fast
check of the app-driving half on its own: about a minute, against the ~25
seconds of OCR plus the same UI work.

    python smoke.py

It leaves a saved order and a saved, paid invoice behind, exactly as a real
run does; nothing here is a dry run.
"""

import fakturama as app
import tracing


#: The document's header, as `extract_sales_order("invoice.pdf")` returns it.
#: Hardcoded so this run exercises the app without paying for OCR first;
#: `main.py` runs the same steps on freshly extracted data.
DEMO_ORDER_DATE = "2026-07-14"
DEMO_EXTERNAL_REFERENCE = "WEB-2026-0714-A17"
DEMO_COMPANY = "Northstar Office GmbH"
DEMO_CONTACT_NAME = "Marta Klein"
DEMO_ALIAS = "NORTHSTAR-BERLIN"
DEMO_PAYMENT_METHOD = "Bank Transfer"
DEMO_PAYMENT_DATE = "2026-07-18"
DEMO_STREET = "Friedrichstrasse 88"
DEMO_POSTCODE = "10117"
DEMO_CITY = "Berlin"
DEMO_COUNTRY = "Germany"
DEMO_EMAIL = "marta.klein@example.test"
DEMO_PHONE = "+49 30 5550 1420"

#: What the document's own totals come to, for the check in step 4.3.
DEMO_TOTALS = {"net": 570.00, "vat": 108.30, "gross": 678.30}

#: The document's item lines, as extraction reads them off invoice.pdf. The
#: gross price and line total are what steps 3.9 and 3.16 work out from the
#: rest; they are spelled out here so the smoke run can check the app's
#: arithmetic without importing the rules it is checking.
DEMO_ITEMS = (
    {
        "sku": "CHR-ERG-01",
        "description": "Ergonomic Desk Chair",
        "qty": 2,
        "unit_price": 250.00,
        "discount_pct": 10,
        "vat_pct": 19,
        "gross_price": 297.50,
        "line_total": 450.00,
    },
    {
        "sku": "MAT-DESK-02",
        "description": "Anti-Fatigue Desk Mat",
        "qty": 3,
        "unit_price": 40.00,
        "discount_pct": 0,
        "vat_pct": 19,
        "gross_price": 47.60,
        "line_total": 120.00,
    },
)


if __name__ == "__main__":
    # Steps 1.3 to 5.7, printing what the app holds after each one.
    tracing.configure(visual=True)
    try:
        window = app.connect()
        print("connected:", window.element_info.name)
        # The item table's right-hand columns are only readable at full width.
        app.maximize(window)

        # --- 1.3-1.7: a new order, with its header filled in ---------------
        order_editor = app.open_new_order(window)
        print("editor:", order_editor.element_info.name)
        app.set_date(order_editor, DEMO_ORDER_DATE)
        print("date:", app.field_value(order_editor, "Date"))
        app.set_customer_reference(order_editor, DEMO_EXTERNAL_REFERENCE)
        print("cust.ref:", app.field_value(order_editor, app.CUSTOMER_REFERENCE_LABEL))
        app.set_price_mode(order_editor)
        app.set_vat_mode(order_editor)
        print("price mode:", app.price_mode(order_editor))
        print("vat:", app.combo_value(order_editor, "VAT"))

        # --- 2.1-2.3: look the debtor up from the order, and decide --------
        first_name, last_name = app.split_contact_name(DEMO_CONTACT_NAME)
        criteria = dict(
            company=DEMO_COMPANY,
            first_name=first_name,
            last_name=last_name,
            postcode=DEMO_POSTCODE,
            city=DEMO_CITY,
        )
        dialog = app.open_address_selector(order_editor, window)
        debtor = app.find_debtor(dialog, DEMO_COMPANY, **criteria)
        print("existing debtor:", debtor.cells if debtor else "none -- create one")

        if not debtor:
            app.dismiss_dialog(dialog, "Cancel")

            # --- 2.5: the debtor editor, beside the still-open order -------
            debtor_editor = app.open_new_debtor(window)
            print("debtor editor:", debtor_editor.element_info.name)

            # 2.6: identity, leaving the proposed Customer ID and "---" alone.
            app.set_debtor_identity(debtor_editor, company=DEMO_COMPANY, contact_name=DEMO_CONTACT_NAME)
            print("customer id:", app.field_value(debtor_editor, "Customer ID"))
            print("company:", app.field_value(debtor_editor, "Company"))
            print("name:", app.field_values(debtor_editor, app.NAME_LABEL))
            print("salutation:", app.combo_value(debtor_editor, "Salutation"))

            # 2.7: the billing address. `main.py` splits the extracted
            # one-liner into these parts with `models.parse_postal_address`.
            app.set_main_address(
                debtor_editor,
                street=DEMO_STREET,
                postcode=DEMO_POSTCODE,
                city=DEMO_CITY,
                country=DEMO_COUNTRY,
                email=DEMO_EMAIL,
                phone=DEMO_PHONE,
            )
            print("street:", app.field_value(debtor_editor, "Street"))
            print("zip/city:", app.field_values(debtor_editor, app.POSTCODE_CITY_LABEL))
            print("country:", app.combo_value(debtor_editor, "Country"))
            print("email:", app.field_value(debtor_editor, "E-Mail"))
            print("phone:", app.field_value(debtor_editor, "Telephone"))

            # 2.8: the invoice address. It would take the delivery role too if
            # the document's two addresses were the same place -- `main.py`
            # decides that with `models.same_address`; this document's differ.
            app.set_address_roles(debtor_editor, invoice=True, delivery=False)
            print("address type:", app.field_value(debtor_editor, app.ADDRESS_TYPE_LABEL))

            # 2.9: alias, no discount, prices net.
            app.set_debtor_miscellaneous(debtor_editor, alias=DEMO_ALIAS)
            print("alias:", app.field_value(debtor_editor, app.ALIAS_LABEL))
            print("discount:", app.field_value(debtor_editor, app.DISCOUNT_LABEL))
            print("net or gross:", app.combo_value(debtor_editor, app.NET_GROSS_LABEL))

            # 2.10: the payment method, creating it if this install lacks it.
            try:
                app.set_debtor_payment(debtor_editor, DEMO_PAYMENT_METHOD)
            except app.PaymentMethodUnavailable as unavailable:
                print("payment: needs creating --", unavailable)

                # 2.10.1-2.10.2: it may exist without being offered here yet.
                if not app.find_payment_method(window, DEMO_PAYMENT_METHOD):
                    # 2.10.3-2.10.5: fill it in. Saving is 2.10.6.
                    term_editor = app.create_payment_method(window, DEMO_PAYMENT_METHOD)
                    print("new term:", app.field_value(term_editor, "Name"))
                    print(
                        "payment code:",
                        app.combo_value(term_editor, app.PAYMENT_CODE_LABEL).strip(),
                    )
                    print(
                        "zeros:",
                        [app.field_value(term_editor, label) for label in (app.CASH_DISCOUNT_LABEL, *app.DAY_LABELS)],
                    )
                    app.save_editor(window, app.NEW_TERM_TAB_RE, "New Term of Payment")

                debtor_editor = app.activate_editor(window, app.NEW_DEBTOR_TAB_RE, "New Debtor")
                app.set_debtor_payment(debtor_editor, DEMO_PAYMENT_METHOD)
            print("payment:", app.combo_value(debtor_editor, app.PAYMENT_LABEL))

            # 2.11: save the debtor, once.
            app.save_editor(window, app.NEW_DEBTOR_TAB_RE, "New Debtor")
            print("debtor saved as:", app.field_value(debtor_editor, "Customer ID"))

            # 2.12: back to the order, and pick the debtor we just saved.
            order_editor = app.activate_editor(window, app.NEW_ORDER_TAB_RE, "New Order")
            dialog = app.open_address_selector(order_editor, window)
            debtor = app.find_debtor(dialog, DEMO_COMPANY, **criteria)
            if not debtor:
                app.dismiss_dialog(dialog, "Cancel")
                raise app.ManualReviewRequired(
                    f"The debtor was saved but does not come back when searching for {DEMO_COMPANY!r}", []
                )

        # --- 2.4 / 2.13: use the debtor, and check what the order got ------
        app.choose_row(dialog, debtor)
        order_editor = app.activate_editor(window, app.NEW_ORDER_TAB_RE, "New Order")
        filled = app.confirm_order_addresses(
            order_editor,
            {app.ROLE_INVOICE: [DEMO_COMPANY, DEMO_STREET, DEMO_POSTCODE, DEMO_CITY, DEMO_COUNTRY]},
        )
        for role, text in filled.items():
            print(f"{role}: {text!r}")
        print("addresses confirmed against the document")

        # --- 3.1-3.17: every item line, in the document's order ------------
        for position, item in enumerate(DEMO_ITEMS, start=1):
            sku, description = item["sku"], item["description"]

            # 3.2-3.3: pick the product from the order, if it is there.
            if not app.select_product(order_editor, window, sku):
                print(f"item {position}: no product {sku!r} -- creating one")

                # 3.4-3.6: the VAT rate first, so the product editor offers it.
                if not app.find_vat(window, item["vat_pct"]):
                    app.create_vat(window, item["vat_pct"])
                    app.save_editor(window, app.NEW_VAT_TAB_RE, "New TAX Rate")
                    print(f"  created the {app.vat_name(item['vat_pct'])} rate")

                # 3.7-3.10: the product itself. `main.py` works the gross
                # price out with `models.gross_price`.
                product_editor = app.create_product(
                    window,
                    sku=sku,
                    description=description,
                    price=item["gross_price"],
                    vat=app.vat_name(item["vat_pct"]),
                )
                print("  item number:", app.field_value(product_editor, "Item Number"))
                print("  price (gross):", app.field_value(product_editor, app.GROSS_PRICE_LABEL))
                print("  vat:", app.combo_value(product_editor, "VAT"))

                # 3.11: save, once.
                app.save_editor(window, app.NEW_PRODUCT_TAB_RE, "New product")

                # 3.12: back to the order, and pick what we just saved.
                order_editor = app.activate_editor(window, app.NEW_ORDER_TAB_RE, "New Order")
                if not app.select_product(order_editor, window, sku):
                    raise app.ManualReviewRequired(
                        f"Product {sku!r} was saved but the picker still does not offer it", []
                    )

            # 3.13-3.15: the quantity and discount this transaction was given.
            order_editor = app.activate_editor(window, app.NEW_ORDER_TAB_RE, "New Order")
            line = app.find_item_line(order_editor, sku)
            if line is None:
                raise app.ManualReviewRequired(f"No order line for {sku!r} to complete", [])
            app.set_item_cell(order_editor, line, app.QTY_COLUMN, f"{item['qty']:g}")
            if item["discount_pct"]:
                app.set_item_cell(order_editor, line, app.LINE_DISCOUNT_COLUMN, f"{item['discount_pct']:g}")

            # 3.14 and 3.16: what the product brought with it, and the total.
            filled_line = app.item_line(order_editor, sku)
            print(f"item {position}: {filled_line.cells if filled_line else 'unreadable'}")
            price = app.money(filled_line.get(app.LINE_PRICE_COLUMN)) if filled_line else None
            if price != item["line_total"]:
                raise app.ManualReviewRequired(
                    f"Line {position} comes to {price!r}, the document says {item['line_total']:.2f}", []
                )

        # --- 4.1-4.3: nothing of the order's own, and totals that agree ----
        app.confirm_order_charges(order_editor)
        print("totals:", app.confirm_order_totals(order_editor, **DEMO_TOTALS))

        # --- 4.4-4.5: save it, then read it back from Data > Documents -----
        number = app.save_order(window, order_editor)
        print("saved as:", number)
        print(
            "documents:",
            app.confirm_document_row(
                window,
                number,
                date=app.shown_date(DEMO_ORDER_DATE),
                reference=DEMO_EXTERNAL_REFERENCE,
                state=app.OPEN_STATE,
                total=f"{DEMO_TOTALS['gross']:.2f}",
            ).cells,
        )

        # --- 4.6-4.7: the invoice, raised from the order to keep the link --
        order_editor = app.activate_editor(window, rf"^\*?{number}$", number)
        invoice_editor = app.create_follow_up(window, order_editor)
        print("invoice editor:", invoice_editor.element_info.name)
        # --- 5.1-5.6: the invoice the order raised, checked and settled ---
        print("invoice carries:", app.confirm_invoice_from_order(window, number))
        invoice_editor = app.activate_editor(window, app.NEW_INVOICE_TAB_RE, "New Invoice")
        app.set_invoice_payment(invoice_editor, DEMO_PAYMENT_METHOD)
        print(
            "payment row:",
            app.set_invoice_paid(
                invoice_editor,
                paid=True,
                payment_date=DEMO_PAYMENT_DATE,
                value=DEMO_TOTALS["gross"],
            ),
        )
        invoice_number = app.save_invoice(window, invoice_editor)
        print("invoice saved as:", invoice_number)
        for document, row in app.confirm_documents(
            window,
            {
                invoice_number: dict(state=app.PAID_STATE, total=f"{DEMO_TOTALS['gross']:.2f}",
                                     reference=DEMO_EXTERNAL_REFERENCE),
                number: dict(state=app.OPEN_STATE, total=f"{DEMO_TOTALS['gross']:.2f}",
                             reference=DEMO_EXTERNAL_REFERENCE),
            },
        ).items():
            print(f"{document}: {row.cells}")
        print(
            "saved invoice holds:",
            app.confirm_saved_invoice(
                window,
                invoice_number,
                method=DEMO_PAYMENT_METHOD,
                paid=True,
                payment_date=DEMO_PAYMENT_DATE,
                value=DEMO_TOTALS["gross"],
            ),
        )

        # 5.7: the flow ends here. No Delivery, Correction or Dunning.
        print(f"done: order {number}, invoice {invoice_number}. Nothing further is created.")
    finally:
        tracing.stop()
