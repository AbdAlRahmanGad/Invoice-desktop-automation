from sales_order import extract_sales_order

if __name__ == '__main__':
    order = extract_sales_order('invoice.pdf')
    print(order.model_dump_json(indent=2))
