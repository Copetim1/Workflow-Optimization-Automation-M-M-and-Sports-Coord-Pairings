# Mentor–Mentee Matching Algorithm README

## Overview

This document explains how mentor–mentee compatibility scores are calculated based on user preferences. Each question category contributes a weighted percentage to an overall match score. The weights vary depending on whether the interaction is **Specific**, **Professional**, or **Casual**.

The goal of this system is to balance intent (what users explicitly want) with softer signals like academic alignment and personality compatibility.

---

## Preference Types

### 1. Specific Requests

Specific requests occur when a mentor and/or mentee explicitly request each other.

| Condition                        | Match Contribution |
| -------------------------------- | ------------------ |
| Both request each other          | **100%**           |
| One requests, the other does not | **30%**            |

Specific requests override most other considerations due to high intent.

---

### 2. Professional Matching

Professional matching emphasizes academic and career alignment.

**Weight Breakdown:**

* **Major Match:** 75%
* **Same Department:** ~30% (supplementary factor)
* **Personality Compatibility:** 25%

**Personality Color Weights (Professional):**

| Color       | Contribution |
| ----------- | ------------ |
| Red         | 0%          |
| Yellow      | 25%          |
| Light Green | 50%          |
| Green       | 75%          |
| Blue        | 100%         |

Professional matching prioritizes structured compatibility, especially shared academic interests.

---

### 3. Casual Matching

Casual matching focuses more on personality and general fit, with reduced emphasis on academics.

**Weight Breakdown:**

* **Major Match:**

  * 25% if the user does **not** care about major
  * 50% if the user **does** care about major
* **Personality Compatibility:**

  * 75% if the user does **not** care about major
  * 50% if the user **does** care about major

**Personality Color Weights (Casual):**

| Color       | Contribution |
| ----------- | ------------ |
| Red         | 0%          |
| Yellow      | 25%          |
| Light Green | 50%          |
| Green       | 75%          |
| Blue        | 100%         |

Casual matching adapts dynamically based on how much importance a user places on academic similarity.

---

## Summary

* **Specific requests** carry the highest priority.
* **Professional matches** emphasize major and department alignment.
* **Casual matches** emphasize personality, with flexible weighting based on user preferences.
* **Personality colors** map to fixed percentage contributions across matching modes.

This modular weighting system allows the algorithm to remain flexible while respecting explicit user intent.

---

## Notes

* Percentages are intended as relative weights and may be normalized during final score computation.
* Department matching is optional and acts as a secondary boost rather than a primary factor.

---

End of README.
