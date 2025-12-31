"""
Tests for TabStop XML generation in format_builder.

These tests verify the XML generation for tab stops used in
TOC styles and other paragraph formatting.
"""

from lxml import etree

from python_docx_redline.constants import WORD_NAMESPACE
from python_docx_redline.format_builder import ParagraphPropertyBuilder
from python_docx_redline.models.style import TabStop


class TestTabStopXmlGeneration:
    """Tests for ParagraphPropertyBuilder.tab_stops_to_element method."""

    def test_single_tab_stop(self) -> None:
        """Test generating XML for a single tab stop."""
        tab_stops = [TabStop(position=6.5, alignment="right", leader="dot")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        # Check element is w:tabs
        assert elem.tag == f"{{{WORD_NAMESPACE}}}tabs"

        # Should have one w:tab child
        tabs = elem.findall(f"{{{WORD_NAMESPACE}}}tab")
        assert len(tabs) == 1

    def test_position_in_twips(self) -> None:
        """Test that position is correctly converted to twips."""
        # 6.5 inches = 9360 twips (6.5 * 1440)
        tab_stops = [TabStop(position=6.5)]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        pos = tab.get(f"{{{WORD_NAMESPACE}}}pos")
        assert pos == "9360"

    def test_one_inch_position(self) -> None:
        """Test that 1 inch position equals 1440 twips."""
        tab_stops = [TabStop(position=1.0)]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        pos = tab.get(f"{{{WORD_NAMESPACE}}}pos")
        assert pos == "1440"

    def test_left_alignment(self) -> None:
        """Test left alignment value."""
        tab_stops = [TabStop(position=1.0, alignment="left")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        val = tab.get(f"{{{WORD_NAMESPACE}}}val")
        assert val == "left"

    def test_right_alignment(self) -> None:
        """Test right alignment value."""
        tab_stops = [TabStop(position=1.0, alignment="right")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        val = tab.get(f"{{{WORD_NAMESPACE}}}val")
        assert val == "right"

    def test_center_alignment(self) -> None:
        """Test center alignment value."""
        tab_stops = [TabStop(position=1.0, alignment="center")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        val = tab.get(f"{{{WORD_NAMESPACE}}}val")
        assert val == "center"

    def test_decimal_alignment(self) -> None:
        """Test decimal alignment value."""
        tab_stops = [TabStop(position=1.0, alignment="decimal")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        val = tab.get(f"{{{WORD_NAMESPACE}}}val")
        assert val == "decimal"

    def test_dot_leader(self) -> None:
        """Test dot leader value."""
        tab_stops = [TabStop(position=1.0, leader="dot")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        leader = tab.get(f"{{{WORD_NAMESPACE}}}leader")
        assert leader == "dot"

    def test_hyphen_leader(self) -> None:
        """Test hyphen leader value."""
        tab_stops = [TabStop(position=1.0, leader="hyphen")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        leader = tab.get(f"{{{WORD_NAMESPACE}}}leader")
        assert leader == "hyphen"

    def test_underscore_leader(self) -> None:
        """Test underscore leader value."""
        tab_stops = [TabStop(position=1.0, leader="underscore")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        leader = tab.get(f"{{{WORD_NAMESPACE}}}leader")
        assert leader == "underscore"

    def test_none_leader_not_added(self) -> None:
        """Test that 'none' leader is not added as attribute."""
        tab_stops = [TabStop(position=1.0, leader="none")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        leader = tab.get(f"{{{WORD_NAMESPACE}}}leader")
        assert leader is None

    def test_default_leader_not_added(self) -> None:
        """Test that default leader (none) is not added."""
        tab_stops = [TabStop(position=1.0)]  # Default leader is "none"
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")
        leader = tab.get(f"{{{WORD_NAMESPACE}}}leader")
        assert leader is None

    def test_multiple_tab_stops(self) -> None:
        """Test generating XML for multiple tab stops."""
        tab_stops = [
            TabStop(position=1.0, alignment="left"),
            TabStop(position=3.0, alignment="center"),
            TabStop(position=6.5, alignment="right", leader="dot"),
        ]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tabs = elem.findall(f"{{{WORD_NAMESPACE}}}tab")
        assert len(tabs) == 3

    def test_multiple_tab_stops_positions(self) -> None:
        """Test that multiple tab stops have correct positions."""
        tab_stops = [
            TabStop(position=1.0),
            TabStop(position=3.0),
            TabStop(position=6.5),
        ]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tabs = elem.findall(f"{{{WORD_NAMESPACE}}}tab")
        positions = [tab.get(f"{{{WORD_NAMESPACE}}}pos") for tab in tabs]
        assert positions == ["1440", "4320", "9360"]

    def test_toc_typical_tab_stop(self) -> None:
        """Test a typical TOC tab stop configuration."""
        # TOC entries typically have right-aligned tab at 6.5 inches with dot leader
        tab_stops = [TabStop(position=6.5, alignment="right", leader="dot")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        tab = elem.find(f"{{{WORD_NAMESPACE}}}tab")

        # Check all attributes
        assert tab.get(f"{{{WORD_NAMESPACE}}}pos") == "9360"
        assert tab.get(f"{{{WORD_NAMESPACE}}}val") == "right"
        assert tab.get(f"{{{WORD_NAMESPACE}}}leader") == "dot"

    def test_valid_xml_structure(self) -> None:
        """Test that generated XML has valid OOXML structure."""
        tab_stops = [TabStop(position=6.5, alignment="right", leader="dot")]
        elem = ParagraphPropertyBuilder.tab_stops_to_element(tab_stops)

        # Verify element structure via lxml API
        assert elem.tag.endswith("}tabs")
        assert len(elem) == 1  # One tab child

        tab = elem[0]
        assert tab.tag.endswith("}tab")
        assert tab.get(f"{{{WORD_NAMESPACE}}}pos") == "9360"
        assert tab.get(f"{{{WORD_NAMESPACE}}}val") == "right"
        assert tab.get(f"{{{WORD_NAMESPACE}}}leader") == "dot"


class TestSetTabStops:
    """Tests for ParagraphPropertyBuilder.set_tab_stops method."""

    def test_adds_tabs_to_empty_ppr(self) -> None:
        """Test adding tab stops to an empty pPr element."""
        ppr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        tab_stops = [TabStop(position=6.5, alignment="right", leader="dot")]

        ParagraphPropertyBuilder.set_tab_stops(ppr, tab_stops)

        tabs = ppr.find(f"{{{WORD_NAMESPACE}}}tabs")
        assert tabs is not None
        assert len(tabs) == 1

    def test_replaces_existing_tabs(self) -> None:
        """Test that existing tabs are replaced."""
        ppr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        # Add existing tabs
        old_tabs = etree.SubElement(ppr, f"{{{WORD_NAMESPACE}}}tabs")
        etree.SubElement(old_tabs, f"{{{WORD_NAMESPACE}}}tab")
        etree.SubElement(old_tabs, f"{{{WORD_NAMESPACE}}}tab")

        # Set new tab stops
        tab_stops = [TabStop(position=6.5, alignment="right", leader="dot")]
        ParagraphPropertyBuilder.set_tab_stops(ppr, tab_stops)

        # Should only have one tabs element with one tab
        tabs_elems = ppr.findall(f"{{{WORD_NAMESPACE}}}tabs")
        assert len(tabs_elems) == 1
        assert len(tabs_elems[0]) == 1

    def test_empty_list_removes_tabs(self) -> None:
        """Test that empty list removes existing tabs."""
        ppr = etree.Element(f"{{{WORD_NAMESPACE}}}pPr")
        # Add existing tabs
        old_tabs = etree.SubElement(ppr, f"{{{WORD_NAMESPACE}}}tabs")
        etree.SubElement(old_tabs, f"{{{WORD_NAMESPACE}}}tab")

        # Set empty tab stops
        ParagraphPropertyBuilder.set_tab_stops(ppr, [])

        # Should have no tabs element
        tabs = ppr.find(f"{{{WORD_NAMESPACE}}}tabs")
        assert tabs is None
