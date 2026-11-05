"""Run with: python -m unittest discover tests"""
import math
import sys
import unittest
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from irrbb.calc import (EARLY_WARNING, LIMIT_BREACH, SOT_BREACH, WITHIN, delta_eve, delta_nii, rate_floor,  # noqa: E402
                        shock_bp, shocked_rates, status_for)


class ShockShapes(unittest.TestCase):
    params = (200, 250, 100)

    def test_parallel_is_flat(self):
        self.assertTrue(np.allclose(shock_bp("Parallel up", [0, 5, 30], *self.params), 200))
        self.assertTrue(np.allclose(shock_bp("Parallel down", [0, 5, 30], *self.params), -200))

    def test_short_shock_starts_at_full_size_and_decays(self):
        self.assertAlmostEqual(float(shock_bp("Short up", 0, *self.params)), 250)
        self.assertAlmostEqual(float(shock_bp("Short up", 4, *self.params)), 250 * math.exp(-1))

    def test_steepener_and_flattener_at_the_ends(self):
        self.assertAlmostEqual(float(shock_bp("Steepener", 0, *self.params)), -0.65 * 250)
        self.assertAlmostEqual(float(shock_bp("Flattener", 0, *self.params)), 0.8 * 250)
        self.assertAlmostEqual(float(shock_bp("Steepener", 1000, *self.params)), 0.9 * 100)
        self.assertAlmostEqual(float(shock_bp("Flattener", 1000, *self.params)), -0.6 * 100)


class Floor(unittest.TestCase):
    def test_floor_values(self):
        self.assertAlmostEqual(float(rate_floor(0)), -0.015)
        self.assertAlmostEqual(float(rate_floor(10)), -0.012)
        self.assertEqual(float(rate_floor(50)), 0.0)
        self.assertEqual(float(rate_floor(80)), 0.0)

    def test_floor_binds_on_a_large_down_shock(self):
        self.assertAlmostEqual(float(shocked_rates(np.array(0.01), -400, 0.0)), -0.015)

    def test_base_rate_below_floor_is_kept(self):
        self.assertAlmostEqual(float(shocked_rates(np.array(-0.02), -100, 0.0)), -0.02)


class Sensitivities(unittest.TestCase):
    def test_long_positive_gap_loses_value_when_rates_rise(self):
        t, gap = np.array([10.0]), np.array([1000.0])
        expected = 1000 * (math.exp(-0.05 * 10) - math.exp(-0.03 * 10))
        self.assertAlmostEqual(delta_eve(gap, t, np.array([0.03]), np.array([0.05])), expected)
        self.assertLess(expected, 0)

    def test_nii_only_counts_buckets_repricing_within_a_year(self):
        t, gap = np.array([0.5, 2.0]), np.array([1000.0, 5000.0])
        base, shocked = np.array([0.02, 0.02]), np.array([0.04, 0.04])
        self.assertAlmostEqual(delta_nii(gap, t, base, shocked), 1000 * 0.02 * 0.5)


class Status(unittest.TestCase):
    def test_bands(self):
        self.assertEqual(status_for(-9.9, 10, 12, 15), WITHIN)
        self.assertEqual(status_for(-11, 10, 12, 15), EARLY_WARNING)
        self.assertEqual(status_for(-13, 10, 12, 15), LIMIT_BREACH)
        self.assertEqual(status_for(-16, 10, 12, 15), SOT_BREACH)

    def test_gains_never_consume_a_limit(self):
        self.assertEqual(status_for(20, 10, 12, 15), WITHIN)


if __name__ == "__main__":
    unittest.main()
