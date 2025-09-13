#!/usr/bin/env python3
"""
Test suite for PowerPoint Tracker application
"""

import os
import sys
import unittest
from unittest.mock import Mock, patch, MagicMock
import tempfile
import shutil

# Add the current directory to the path so we can import our modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from presentation_detector import PowerPointTracker, PowerPointWindowDetector, WindowInfo

class TestPowerPointTracker(unittest.TestCase):
    """Test cases for PowerPointTracker class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.temp_dir = tempfile.mkdtemp()
        self.test_ppt_path = os.path.join(self.temp_dir, "test.pptx")
        
        # Create a mock PowerPoint file (we'll mock the actual loading)
        with open(self.test_ppt_path, 'w') as f:
            f.write("Mock PowerPoint file")
    
    def tearDown(self):
        """Clean up test fixtures"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_tracker_initialization_without_file(self):
        """Test tracker initialization without a file"""
        tracker = PowerPointTracker()
        self.assertIsNone(tracker.ppt_path)
        self.assertIsNone(tracker.presentation)
        self.assertEqual(tracker.current_slide_index, 0)
        self.assertEqual(tracker.total_slides, 0)
    
    def test_tracker_initialization_with_auto_detect(self):
        """Test tracker initialization with auto-detect enabled"""
        tracker = PowerPointTracker(auto_detect=True)
        self.assertTrue(tracker.auto_detect)
        self.assertIsNotNone(tracker.window_detector)
    
    def test_tracker_initialization_with_file(self):
        """Test tracker initialization with a file path"""
        with patch('pptx.Presentation') as mock_presentation:
            mock_slides = [Mock(), Mock(), Mock()]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            
            self.assertEqual(tracker.ppt_path, self.test_ppt_path)
            self.assertEqual(tracker.total_slides, 3)
            self.assertEqual(tracker.current_slide_index, 0)
    
    def test_load_presentation_file_not_found(self):
        """Test loading a presentation file that doesn't exist"""
        tracker = PowerPointTracker()
        
        with self.assertRaises(FileNotFoundError):
            tracker.load_presentation()
    
    def test_slide_navigation(self):
        """Test slide navigation methods"""
        with patch('pptx.Presentation') as mock_presentation:
            mock_slides = [Mock(), Mock(), Mock()]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            
            # Test next slide
            self.assertTrue(tracker.next_slide())
            self.assertEqual(tracker.current_slide_index, 1)
            
            # Test next slide at end
            self.assertTrue(tracker.next_slide())
            self.assertEqual(tracker.current_slide_index, 2)
            
            # Test next slide beyond end
            self.assertFalse(tracker.next_slide())
            self.assertEqual(tracker.current_slide_index, 2)
            
            # Test previous slide
            self.assertTrue(tracker.previous_slide())
            self.assertEqual(tracker.current_slide_index, 1)
            
            # Test previous slide at beginning
            self.assertTrue(tracker.previous_slide())
            self.assertEqual(tracker.current_slide_index, 0)
            
            # Test previous slide beyond beginning
            self.assertFalse(tracker.previous_slide())
            self.assertEqual(tracker.current_slide_index, 0)
    
    def test_go_to_slide(self):
        """Test jumping to specific slides"""
        with patch('pptx.Presentation') as mock_presentation:
            mock_slides = [Mock(), Mock(), Mock()]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            
            # Test valid slide numbers
            self.assertTrue(tracker.go_to_slide(1))
            self.assertEqual(tracker.current_slide_index, 0)
            
            self.assertTrue(tracker.go_to_slide(2))
            self.assertEqual(tracker.current_slide_index, 1)
            
            self.assertTrue(tracker.go_to_slide(3))
            self.assertEqual(tracker.current_slide_index, 2)
            
            # Test invalid slide numbers
            self.assertFalse(tracker.go_to_slide(0))
            self.assertFalse(tracker.go_to_slide(4))
            self.assertFalse(tracker.go_to_slide(-1))
    
    def test_get_current_slide_number(self):
        """Test getting current slide number"""
        with patch('pptx.Presentation') as mock_presentation:
            mock_slides = [Mock(), Mock(), Mock()]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            
            self.assertEqual(tracker.get_current_slide_number(), 1)
            
            tracker.next_slide()
            self.assertEqual(tracker.get_current_slide_number(), 2)
            
            tracker.next_slide()
            self.assertEqual(tracker.get_current_slide_number(), 3)
    
    def test_get_total_slides(self):
        """Test getting total number of slides"""
        with patch('pptx.Presentation') as mock_presentation:
            mock_slides = [Mock(), Mock(), Mock()]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            self.assertEqual(tracker.get_total_slides(), 3)
    
    def test_extract_slide_text(self):
        """Test text extraction from slides"""
        with patch('pptx.Presentation') as mock_presentation:
            # Create mock slide with text
            mock_shape = Mock()
            mock_shape.text = "Test slide text"
            mock_slide = Mock()
            mock_slide.shapes = [mock_shape]
            
            mock_slides = [mock_slide]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            text = tracker.extract_slide_text()
            
            self.assertEqual(text, "Test slide text")
    
    def test_search_text_in_slides(self):
        """Test searching for text across slides"""
        with patch('pptx.Presentation') as mock_presentation:
            # Create mock slides with different text
            mock_shape1 = Mock()
            mock_shape1.text = "First slide with keyword"
            mock_slide1 = Mock()
            mock_slide1.shapes = [mock_shape1]
            
            mock_shape2 = Mock()
            mock_shape2.text = "Second slide without keyword"
            mock_slide2 = Mock()
            mock_slide2.shapes = [mock_shape2]
            
            mock_shape3 = Mock()
            mock_shape3.text = "Third slide with keyword again"
            mock_slide3 = Mock()
            mock_slide3.shapes = [mock_shape3]
            
            mock_slides = [mock_slide1, mock_slide2, mock_slide3]
            mock_presentation.return_value.slides = mock_slides
            
            tracker = PowerPointTracker(self.test_ppt_path)
            
            # Mock OCR to return empty strings
            with patch.object(tracker, 'extract_text_with_ocr', return_value=""):
                results = tracker.search_text_in_slides("keyword")
                self.assertEqual(results, [1, 3])  # Slides 1 and 3 contain "keyword"
    
    def test_auto_sync_functionality(self):
        """Test auto-sync functionality"""
        tracker = PowerPointTracker(auto_detect=True)
        
        # Test enabling auto-sync
        self.assertTrue(tracker.enable_auto_sync())
        self.assertTrue(tracker.auto_sync_enabled)
        
        # Test disabling auto-sync
        tracker.disable_auto_sync()
        self.assertFalse(tracker.auto_sync_enabled)
    
    def test_auto_sync_without_detector(self):
        """Test auto-sync without window detector"""
        tracker = PowerPointTracker(auto_detect=False)
        
        # Should return False when no detector is available
        self.assertFalse(tracker.enable_auto_sync())
        self.assertFalse(tracker.auto_sync_enabled)

class TestPowerPointWindowDetector(unittest.TestCase):
    """Test cases for PowerPointWindowDetector class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.detector = PowerPointWindowDetector()
    
    def test_detector_initialization(self):
        """Test detector initialization"""
        self.assertIsNotNone(self.detector.system)
        self.assertIsInstance(self.detector.powerpoint_process_names, list)
        self.assertIn("Microsoft PowerPoint", self.detector.powerpoint_process_names)
    
    def test_is_powerpoint_window(self):
        """Test PowerPoint window identification"""
        # Test positive cases
        self.assertTrue(self.detector.is_powerpoint_window("Presentation.pptx - Microsoft PowerPoint", "Microsoft PowerPoint"))
        self.assertTrue(self.detector.is_powerpoint_window("Slide 1 of 10", "PowerPoint"))
        self.assertTrue(self.detector.is_powerpoint_window("My Presentation.ppt", "POWERPNT.EXE"))
        
        # Test negative cases
        self.assertFalse(self.detector.is_powerpoint_window("Notepad", "notepad.exe"))
        self.assertFalse(self.detector.is_powerpoint_window("Chrome Browser", "chrome.exe"))
    
    def test_extract_slide_info_from_title(self):
        """Test slide information extraction from window titles"""
        # Test various title formats
        test_cases = [
            ("Slide 1 of 10", {"current_slide": 1, "total_slides": 10}),
            ("Presentation.pptx - Slide 5 of 15", {"current_slide": 5, "total_slides": 15}),
            ("Slide Show - Slide 3 of 8", {"current_slide": 3, "total_slides": 8, "mode": "slideshow"}),
            ("My Presentation.pptx - Slide 2 of 5", {"current_slide": 2, "total_slides": 5, "presentation_name": "My Presentation.pptx"}),
        ]
        
        for title, expected in test_cases:
            result = self.detector.extract_slide_info_from_title(title)
            for key, value in expected.items():
                self.assertEqual(result.get(key), value, f"Failed for title: {title}")
    
    def test_window_info_creation(self):
        """Test WindowInfo object creation"""
        window_info = WindowInfo(
            window_id=12345,
            title="Test Window",
            app_name="Test App",
            position=(100, 200),
            size=(800, 600)
        )
        
        self.assertEqual(window_info.window_id, 12345)
        self.assertEqual(window_info.title, "Test Window")
        self.assertEqual(window_info.app_name, "Test App")
        self.assertEqual(window_info.position, (100, 200))
        self.assertEqual(window_info.size, (800, 600))
        
        # Test string representation
        str_repr = str(window_info)
        self.assertIn("12345", str_repr)
        self.assertIn("Test Window", str_repr)
        self.assertIn("Test App", str_repr)

class TestIntegration(unittest.TestCase):
    """Integration tests"""
    
    def test_tracker_with_detector_integration(self):
        """Test integration between tracker and detector"""
        with patch('window_detector.PowerPointWindowDetector.get_powerpoint_windows') as mock_get_windows:
            # Mock window detection
            mock_window = WindowInfo(
                window_id=12345,
                title="Test Presentation.pptx - Slide 2 of 5",
                app_name="Microsoft PowerPoint"
            )
            mock_get_windows.return_value = [mock_window]
            
            tracker = PowerPointTracker(auto_detect=True)
            window_info = tracker.get_window_info()
            
            self.assertIsNotNone(window_info)
            self.assertEqual(window_info['window_title'], "Test Presentation.pptx - Slide 2 of 5")
            self.assertEqual(window_info['app_name'], "Microsoft PowerPoint")
            
            slide_info = window_info['slide_info']
            self.assertEqual(slide_info['current_slide'], 2)
            self.assertEqual(slide_info['total_slides'], 5)

def run_tests():
    """Run all tests"""
    print("🧪 Running PowerPoint Tracker Test Suite")
    print("=" * 50)
    
    # Create test suite
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add test cases
    suite.addTests(loader.loadTestsFromTestCase(TestPowerPointTracker))
    suite.addTests(loader.loadTestsFromTestCase(TestPowerPointWindowDetector))
    suite.addTests(loader.loadTestsFromTestCase(TestIntegration))
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "=" * 50)
    print("📊 Test Summary:")
    print(f"   Tests run: {result.testsRun}")
    print(f"   Failures: {len(result.failures)}")
    print(f"   Errors: {len(result.errors)}")
    
    if result.failures:
        print("\n❌ Failures:")
        for test, traceback in result.failures:
            print(f"   {test}: {traceback}")
    
    if result.errors:
        print("\n❌ Errors:")
        for test, traceback in result.errors:
            print(f"   {test}: {traceback}")
    
    if result.wasSuccessful():
        print("\n✅ All tests passed!")
        return True
    else:
        print("\n❌ Some tests failed!")
        return False

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)
