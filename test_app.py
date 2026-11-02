import unittest
from app import app
from database import get_connection

class InventoryAppTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.client.testing = True

    def login_as(self, username, password):
        return self.client.post('/login', data={'username': username, 'password': password}, follow_redirects=True)

    def test_login_and_logout(self):
        res = self.login_as('admin', 'admin123')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Shop Overview', res.data)

        # Logout
        res_logout = self.client.get('/logout', follow_redirects=True)
        self.assertEqual(res_logout.status_code, 200)
        self.assertIn(b'Staff Sign In', res_logout.data)

    def test_dashboard_admin(self):
        self.login_as('admin', 'admin123')
        res = self.client.get('/')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Shop Overview', res.data)

    def test_salesman_role_restrictions(self):
        # Login as Salesman
        self.login_as('sales1', 'sales123')

        # Salesman can access POS
        res_pos = self.client.get('/pos')
        self.assertEqual(res_pos.status_code, 200)

        # Salesman CANNOT access settings (redirects)
        res_settings = self.client.get('/settings', follow_redirects=True)
        self.assertIn(b'Access Restricted', res_settings.data)

        # Salesman CANNOT access purchases / stock in
        res_pur = self.client.get('/purchases', follow_redirects=True)
        self.assertIn(b'Access Restricted', res_pur.data)

    def test_warehouse_role_access(self):
        # Login as warehouse
        self.login_as('godown', 'godown123')

        # Can access stock in
        res_pur = self.client.get('/purchases')
        self.assertEqual(res_pur.status_code, 200)

        # Cannot access accounting khata ledger
        res_ledger = self.client.get('/ledger', follow_redirects=True)
        self.assertIn(b'Access Restricted', res_ledger.data)

    def test_accounting_role_access(self):
        # Login as accountant
        self.login_as('accounts', 'accounts123')

        # Can access wholesale khata ledger
        res_ledger = self.client.get('/ledger')
        self.assertEqual(res_ledger.status_code, 200)
        self.assertIn(b'Wholesale Credit Ledger', res_ledger.data)

    def test_inventory_page(self):
        self.login_as('admin', 'admin123')
        res = self.client.get('/inventory')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Products & Stock Catalog', res.data)

    def test_pos_page(self):
        self.login_as('admin', 'admin123')
        res = self.client.get('/pos')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Billing Mode:', res.data)

    def test_purchases_page(self):
        self.login_as('admin', 'admin123')
        res = self.client.get('/purchases')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Stock In', res.data)

    def test_sales_page(self):
        self.login_as('admin', 'admin123')
        res = self.client.get('/sales')
        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Sales History', res.data)

    def test_sale_creation_and_stock_deduction(self):
        self.login_as('sales1', 'sales123')
        conn = get_connection()
        prod = conn.execute("SELECT id, stock_qty, retail_price FROM products WHERE stock_qty > 2 LIMIT 1").fetchone()
        conn.close()
        self.assertIsNotNone(prod)
        prod_id = prod['id']
        initial_stock = prod['stock_qty']

        payload = {
            'sale_type': 'RETAIL',
            'customer_name': 'Multi-User Customer',
            'customer_phone': '9876543210',
            'subtotal': prod['retail_price'] * 2,
            'discount': 0,
            'grand_total': prod['retail_price'] * 2,
            'payment_mode': 'CASH',
            'paid_amount': prod['retail_price'] * 2,
            'items': [
                {
                    'id': prod_id,
                    'quantity': 2,
                    'price': prod['retail_price']
                }
            ]
        }

        res = self.client.post('/api/sales/create', json=payload)
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertTrue(data['success'])
        sale_id = data['sale_id']

        # Verify stock decreased by 2
        conn = get_connection()
        updated_prod = conn.execute("SELECT stock_qty FROM products WHERE id = ?", (prod_id,)).fetchone()
        conn.close()
        self.assertEqual(updated_prod['stock_qty'], initial_stock - 2)

        # Verify print invoice route
        res_inv = self.client.get(f'/invoice/{sale_id}')
        self.assertEqual(res_inv.status_code, 200)
        self.assertIn(b'Multi-User Customer', res_inv.data)

    def test_export_csv(self):
        self.login_as('accounts', 'accounts123')
        res = self.client.get('/export/inventory/csv')
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.mimetype, 'text/csv')

if __name__ == '__main__':
    unittest.main()
