A MS word web and native addin for pentest report writing. 

# Features

## Sort findings accoring to severity

Allow user to mark a section. Sort the subsections according to their severity (Informational, Low, Medium, High, Critical). The severity can be found in a sub-sub section.

example (in markdown)

```

# Background

lorum ipsum

# Findings

## XSS in Y

### Risk: Medium (5.4)

## SQLi

### Risk: Critical (9.0)

## SSRF A

### Risk: Medium (4.0)

```


Should reorder the findings sections to:

```

# Background

lorum ipsum

# Findings

## SQLi

### Risk: Critical (9.0)

## XSS in Y

### Risk: Medium (5.4)


## SSRF A

### Risk: Medium (4.0)

```

## insert rainbow findings table

User selects a section from a list. Then create a table of the subsection headings. The table should have the following columns:

* # (heading number)
* Severity (informational, low, medium, high, critical)
* Score (if present in the subsections)
* Title

Severity and score can be read from subsections of each finding.


# Design Goals

Follow KISS and YAGNI.

The addon will be distributed/shared with a small number of people through onedrive/sharepoint.


