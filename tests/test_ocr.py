"""OCR engine tests.

The regressions that matter here are real ones: v1.0.1 inverted every capture and cut it at a
fixed threshold, which silently garbled or erased any text lighter than about #5a5a5a.
"""

import unittest

from helpers import (SENTENCE, load_quickocr, normalise, render_text, require_tesseract,
                     require_windows)

quickocr = load_quickocr()


class OCRTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        require_windows()
        require_tesseract()
        cls.engine = quickocr.OCREngine()

    def assertReads(self, image, expected=SENTENCE, msg=None):
        result = self.engine.extract_text(image)
        self.assertEqual(normalise(result.text), normalise(expected),
                         msg or f"variant={result.variant} conf={result.confidence:.0f}")
        return result


class TestReadsRealWorldText(OCRTestCase):
    def test_black_on_white(self):
        self.assertReads(render_text())

    def test_light_on_dark(self):
        self.assertReads(render_text(fg=(235, 235, 235), bg=(20, 20, 20)))

    def test_coloured_on_dark(self):
        self.assertReads(render_text(fg=(200, 164, 100), bg=(58, 47, 30)))

    def test_low_contrast_grey_text(self):
        """v1.0.1 turned these into garbage or an entirely blank frame."""
        for level in (90, 110, 130):
            with self.subTest(grey=level):
                self.assertReads(render_text(fg=(level,) * 3))

    def test_grey_on_off_white(self):
        self.assertReads(render_text(fg=(136, 136, 136), bg=(240, 240, 240)))


class TestVariantSelection(OCRTestCase):
    def test_noisy_artwork_uses_the_fixed_threshold(self):
        """Otsu splits the background when text is only a few percent of the pixels, so the
        fixed cut has to stay available as a candidate."""
        result = self.engine.extract_text(
            render_text(fg=(200, 164, 100), bg=(58, 47, 30), noise=40))
        self.assertEqual(result.variant, 'fixed')
        self.assertGreater(result.confidence, 70)
        self.assertIn("1,024.50", result.text)

    def test_clean_text_stops_after_the_first_variant(self):
        result = self.engine.extract_text(render_text())
        self.assertEqual(result.variant, 'otsu')
        self.assertGreaterEqual(result.confidence, quickocr.CONF_ACCEPT)

    def test_blank_image_reports_no_text_and_no_confidence(self):
        from PIL import Image
        result = self.engine.extract_text(Image.new("RGB", (300, 60), (255, 255, 255)))
        self.assertEqual(result.text.strip(), "")
        self.assertEqual(result.confidence, 0.0)


class TestUnicode(OCRTestCase):
    def test_installed_language_characters_survive(self):
        """Not an encoding test of Tesseract's accuracy - a check that non-ASCII passes
        through the pipeline intact for a language that is actually installed."""
        if not self.engine.supports('fra'):
            self.skipTest("French language data is not installed")
        engine = quickocr.OCREngine('fra')
        result = engine.extract_text(render_text("hôtel très élégant",
                                                 width=300, size=20))
        self.assertIn("ô", result.text)
        self.assertIn("è", result.text)

    def test_symbols_pass_through(self):
        result = self.engine.extract_text(
            render_text("Prix 1 234,56 € «guillemets»", width=420, size=20))
        self.assertIn("€", result.text)


class TestLanguages(OCRTestCase):
    def test_discovers_installed_languages(self):
        self.assertIn('eng', self.engine.available)
        self.assertNotIn('osd', self.engine.available,
                         "osd is orientation data, not a language")

    def test_supports_only_reports_installed_languages(self):
        self.assertTrue(self.engine.supports('eng'))
        self.assertFalse(self.engine.supports('zzz'))
        self.assertFalse(self.engine.supports(''))

    def test_choices_lead_with_the_default_combination(self):
        choices = self.engine.language_choices()
        if self.engine.supports(quickocr.DEFAULT_LANG):
            self.assertEqual(choices[0], quickocr.DEFAULT_LANG)
        self.assertEqual(len(choices), len(set(choices)), "choices must not repeat")

    def test_falls_back_when_a_saved_language_is_missing(self):
        engine = quickocr.OCREngine('deu-not-installed')
        self.assertIn(engine.lang, engine.available)


if __name__ == '__main__':
    unittest.main()
