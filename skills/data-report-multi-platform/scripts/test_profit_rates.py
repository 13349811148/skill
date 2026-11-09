import unittest
from profit_rates import net_margin
from store_schema import STORE_HEADERS, compatible_store_headers


class ProfitRateTests(unittest.TestCase):
    def test_net_sales_denominator_preserves_loss_and_confirmed_zero(self):
        self.assertAlmostEqual(net_margin(30,150),.2)
        self.assertAlmostEqual(net_margin(-12,80),-.15)
        self.assertEqual(net_margin(0,80),0)

    def test_missing_nonpositive_denominator_does_not_invent_a_rate(self):
        for profit,sales in [(None,100),('',100),(10,None),(10,''),(10,0),(10,-20),(-10,0)]:
            self.assertIsNone(net_margin(profit,sales))
        with self.assertRaises(ValueError):net_margin(float('nan'),100)

    def test_known_store_schema_migration_preserves_original_columns(self):
        self.assertEqual(STORE_HEADERS[-2:],['店铺净利润','净利率'])
        for count in [26,27,28]:self.assertTrue(compatible_store_headers(STORE_HEADERS[:count]))
        self.assertFalse(compatible_store_headers(STORE_HEADERS+['人工新增']))
        self.assertFalse(compatible_store_headers(['日期','净利率']))


if __name__=='__main__':unittest.main()
