// Fully synthetic fixture for export-typst-to-docx.
// It is intentionally unrelated to any real thesis, experiment, or dataset.

= Synthetic Signal Processing Study

This chapter demonstrates a small fictional processing pipeline. Its names,
values, equations, and diagram exist only to exercise the converter.

== Overview

The synthetic system transforms an input sequence into an output sequence.
As shown in @fig:synthetic-pipeline, the calculation uses the parameters in
@tab:synthetic-parameters and follows @eqt:synthetic-linear-model.

#capfig(
  image("figures/synthetic-pipeline.svg", width: 78%),
  caption: [Synthetic processing pipeline],
  label: <fig:synthetic-pipeline>,
)

== Method

=== Linear response with a hidden Typst label <sec:synthetic-linear-response>

The editable display equation is

$
y(t) = a x(t) + b
$ <synthetic-linear-model>

$10$ samples are used only to confirm that a paragraph beginning with inline
math remains ordinary prose rather than becoming a display equation.

#captab(
  caption: [Synthetic parameter values],
  label: <tab:synthetic-parameters>,
)[
| Parameter | Value | Description |
| --- | --- | --- |
| Gain | $1.25$ | Dimensionless invented value |
| Offset | $0.50$ | Dimensionless invented value |
| Samples | $10$ | Small smoke-test count |
]

== Summary

The fixture contains one figure, one native table, one numbered display
equation, inline formulas, and resolvable figure, table, and equation
cross-references.
