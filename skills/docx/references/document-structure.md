# Document Structure Frameworks

Structure transforms a collection of paragraphs into a persuasive argument. These frameworks ensure your document builds logically toward your conclusion.

## Table of Contents

1. [The Pyramid Principle](#the-pyramid-principle)
2. [SCQA Framework](#scqa-framework)
3. [IRAC for Legal Documents](#irac-for-legal-documents)
4. [Problem-Solution-Benefit](#problem-solution-benefit)
5. [The One-Page Outline Method](#the-one-page-outline-method)
6. [Framework Selection Guide](#framework-selection-guide)

---

## The Pyramid Principle

Barbara Minto's framework, developed at McKinsey: **Lead with the answer, then support it.**

### Core Concept

```
                    ┌─────────────────┐
                    │   MAIN POINT    │
                    │   (Answer)      │
                    └────────┬────────┘
           ┌─────────────────┼─────────────────┐
           ▼                 ▼                 ▼
    ┌─────────────┐   ┌─────────────┐   ┌─────────────┐
    │ Supporting  │   │ Supporting  │   │ Supporting  │
    │  Argument 1 │   │  Argument 2 │   │  Argument 3 │
    └──────┬──────┘   └──────┬──────┘   └──────┬──────┘
           │                 │                 │
      ┌────┴────┐       ┌────┴────┐       ┌────┴────┐
      ▼    ▼    ▼       ▼    ▼    ▼       ▼    ▼    ▼
    Data Data Data    Data Data Data    Data Data Data
```

### The Three Rules

1. **Start with the answer** - Don't make executives wait
2. **Group and summarize** - Ideas at any level summarize the ideas below them
3. **Logic within groups** - Ideas in each group are logically ordered

### MECE Requirement

Supporting arguments must be:
- **Mutually Exclusive** - No overlap between categories
- **Collectively Exhaustive** - Nothing important left out

**Bad (Not MECE):**
- Increase revenue
- Cut costs
- Improve marketing ← overlaps with revenue
- Reduce headcount ← overlaps with costs

**Good (MECE):**
- Increase revenue (3 initiatives)
- Reduce costs (2 initiatives)
- Improve capital efficiency (1 initiative)

### Logical Ordering

Ideas within groups should follow one of these orders:

| Order Type | When to Use | Example |
|------------|-------------|---------|
| **Time** | Process, history, sequence | First... Then... Finally... |
| **Structure** | Geography, organization | North... South... East... West... |
| **Importance** | Recommendations | Most critical... Important... Nice to have... |
| **Deductive** | Logical argument | If X and Y, then Z |

### Applying the Pyramid to Documents

**Traditional (Burying the Lead):**
```
1. Background
2. Methodology
3. Findings
4. Analysis
5. Conclusion  ← Reader finally gets the answer on page 20
```

**Pyramid Structure:**
```
1. Executive Summary
   - Recommendation stated immediately
   - Three supporting reasons
   - Next steps

2. Supporting Argument 1
   - Key point restated
   - Evidence

3. Supporting Argument 2
   - Key point restated
   - Evidence

4. Supporting Argument 3
   - Key point restated
   - Evidence

5. Appendix
   - Detailed methodology
   - Data tables
```

### Example Application

**Question:** Should we expand into the European market?

**Pyramid answer:**

```
RECOMMENDATION: Yes, expand to Europe starting with UK and Germany

Supporting Argument 1: Market opportunity is substantial
  - €2B addressable market in target countries
  - Growing 15% annually
  - Competitors have <20% share

Supporting Argument 2: We have competitive advantages
  - Product already localized for EU regulations
  - Two existing enterprise customers in UK
  - Partnership with regional distributor

Supporting Argument 3: Risk is manageable
  - Capital requirement of €5M within budget
  - Existing team can manage expansion
  - Exit strategy if unsuccessful

NEXT STEPS: Approve €5M budget, begin hiring in Q1
```

---

## SCQA Framework

The McKinsey standard for structuring business arguments.

### Components

| Element | Purpose | Length |
|---------|---------|--------|
| **Situation** | Establish shared context | 1-2 paragraphs |
| **Complication** | Introduce tension/problem | 1-2 paragraphs |
| **Question** | Frame what must be answered | Often implicit |
| **Answer** | Your recommendation | Rest of document |

### Example

**Situation:**
> GlobalTech is the #2 player in enterprise software with 25% market share and $500M annual revenue. For five years, growth has averaged 12% annually, outpacing the market.

**Complication:**
> Over the past 18 months, three cloud-native competitors have captured 15% of new deals. Win rates against these competitors have dropped from 60% to 35%. At current trajectory, GlobalTech will lose market position within three years.

**Question (explicit or implicit):**
> How should GlobalTech respond to protect and grow market share?

**Answer:**
> Accelerate cloud migration with a $50M investment over 18 months. This investment will (1) modernize the core platform, (2) enable competitive pricing, and (3) position for next-generation features.

### SCQA in Document Structure

```
EXECUTIVE SUMMARY
├── Situation (paragraph 1)
├── Complication (paragraph 2)
├── Recommendation (paragraph 3)
└── Key supporting points (bullets)

BODY
├── Detailed Situation Analysis
├── Complication Deep Dive
├── Recommendation Detail
│   ├── Approach 1
│   ├── Approach 2
│   └── Approach 3
├── Implementation Plan
└── Financial Impact

APPENDIX
├── Methodology
├── Data Sources
└── Detailed Analysis
```

### When to Use SCQA

- Executive presentations
- Board recommendations
- Strategic decisions
- Investment cases
- Any document requiring a clear recommendation

---

## IRAC for Legal Documents

The standard framework for legal analysis.

### Components

| Element | Purpose |
|---------|---------|
| **Issue** | State the legal question |
| **Rule** | State the applicable law |
| **Analysis** | Apply the law to the facts |
| **Conclusion** | Answer the question |

### CRAC Variation

**CRAC** (Conclusion-Rule-Analysis-Conclusion) is often preferred in practice because it leads with the answer:

```
CONCLUSION: The contract is likely enforceable.

RULE: Under New York law, a contract requires offer, acceptance,
consideration, and mutual assent. Smith v. Jones, 123 N.Y.2d 456 (2020).

ANALYSIS: Here, Seller made a written offer on January 15, Buyer
accepted in writing on January 20, and both parties exchanged value...

CONCLUSION: Because all elements are satisfied, the contract is
likely enforceable. However, Buyer should be aware of the statute
of frauds issue regarding the amendment.
```

### CRRACC for Complex Issues

For issues with counterarguments:

```
Conclusion
Rule statement
Rule explanation
Application
Counterargument
Conclusion (restated)
```

### Legal Memo Structure with IRAC

```
MEMORANDUM

TO: Senior Partner
FROM: Associate
RE: Client Matter - Contract Enforceability
DATE: January 15, 2025

QUESTION PRESENTED
Whether the January 2024 supply agreement between Client and Vendor
is enforceable under New York law, given the alleged oral modification.

BRIEF ANSWER
Probably yes. The original agreement is enforceable, and while the
oral modification raises statute of frauds concerns, the partial
performance exception likely applies.

STATEMENT OF FACTS
[Neutral recitation of relevant facts]

DISCUSSION

I. The Original Agreement Is Enforceable
   A. [Issue 1 - IRAC analysis]
   B. [Issue 2 - IRAC analysis]

II. The Oral Modification Raises Questions
   A. The Statute of Frauds Generally Applies
   B. The Partial Performance Exception Likely Saves the Modification

CONCLUSION
[Summary of conclusions and recommendation]
```

---

## Problem-Solution-Benefit

The classic persuasive structure, simpler than SCQA for straightforward proposals.

### Components

| Element | Purpose | Typical Length |
|---------|---------|----------------|
| **Problem** | Create urgency | 1-2 pages |
| **Solution** | Your proposal | 2-5 pages |
| **Benefit** | Why it matters | 1-2 pages |

### Example Structure

```
1. THE PROBLEM

   Current customer onboarding takes 45 days on average, causing:
   • 30% of signed customers to churn before activation
   • $2M annual revenue loss from delayed deployments
   • Sales team frustration and reduced pipeline

2. THE SOLUTION

   Implement automated onboarding platform:
   • Self-service configuration wizard
   • Automated data migration
   • Real-time progress tracking

3. THE BENEFIT

   Expected results within 6 months:
   • Reduce onboarding time from 45 days to 14 days
   • Reduce pre-activation churn by 80%
   • Recover $1.6M in annual revenue
   • ROI of 3x on implementation cost

4. NEXT STEPS

   Request approval for $200K implementation budget and 2 FTE
   allocation for 6-month project.
```

### When to Use

- Product proposals
- Process improvement requests
- Budget requests
- Internal business cases
- Customer proposals

---

## The One-Page Outline Method

Build the structure before any content. Similar to the "ghost deck" method for presentations.

### Process

1. **Write section headings only** - No content, just structure
2. **Arrange logically** - Does the document flow?
3. **Test the narrative** - Read headings in sequence
4. **Iterate** - Adjust until headings tell complete story
5. **Fill content** - Only now write body text

### Example: Strategy Recommendation

**Initial outline (headings only):**

```
1. We should acquire TechCo
2. Market dynamics favor consolidation
3. TechCo fills our product gap
4. Valuation is reasonable
5. Integration risk is manageable
6. Request: Approve LOI
```

Reading just these headings tells the complete story.

### The Readability Test

If someone reads only your section headings and understands:
- What the situation is
- What you're recommending
- Why they should agree
- What you want them to do

Then your outline works.

### Common Outline Problems

| Problem | Symptom | Fix |
|---------|---------|-----|
| Topic headings | "Background", "Analysis" | Convert to action headings |
| Buried lead | Recommendation on page 15 | Move to page 1 |
| Missing logic | Jump from problem to request | Add supporting sections |
| Too many points | 20 sections | Consolidate to 3-5 themes |

---

## Framework Selection Guide

| Scenario | Best Framework | Why |
|----------|---------------|-----|
| Board recommendation | SCQA + Pyramid | Clear decision structure |
| Complex analysis | Pyramid | Handles multiple arguments |
| Legal question | IRAC/CRAC | Standard legal reasoning |
| Proposal/pitch | Problem-Solution-Benefit | Drives action |
| Any document | One-Page Outline | Universal planning tool |

### Combining Frameworks

Frameworks can be combined:

- **SCQA + Pyramid:** Use SCQA for overall arc, Pyramid for the Answer section
- **Problem-Solution-Benefit within SCQA:** PSB becomes the Answer section
- **One-Page Outline + Any:** Always outline first, then apply chosen framework

---

## Quick Reference

### Pyramid Checklist

- [ ] Main recommendation stated in first paragraph
- [ ] 3-5 supporting arguments, no more
- [ ] Each argument group is MECE
- [ ] Each argument has supporting evidence
- [ ] No redundancy between sections

### SCQA Checklist

- [ ] Situation: context everyone agrees on
- [ ] Complication: tension or problem that changed things
- [ ] Question: what decision must be made (explicit or implicit)
- [ ] Answer: clear recommendation stated early

### Document Flow Test

Read only your section headings in sequence. Can someone understand:
1. What you're recommending?
2. Why they should agree?
3. What you want them to do?

If not, restructure before writing.

### The "So What?" Test

After every section, ask: "So what?"

If you can't articulate why this section matters to your main argument, either:
1. Clarify the connection explicitly
2. Move to appendix
3. Delete it
