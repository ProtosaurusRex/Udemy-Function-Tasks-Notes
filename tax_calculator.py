
def tax_calculator(subtotal):   
    tax_totals = []
    tax_totals.append(subtotal)
    tax_amount = round(subtotal * 0.08, 2)
    tax_totals.append(tax_amount)
    total = subtotal + tax_amount
    tax_totals.append(total)
    return tax_totals
