= Signal mapping

#let formula-note(body) = block(width: 100%)[
  #set par(first-line-indent: (amount: 0pt, all: true))
  式中，#body
]

The fictional model in @eqt:chapter-two-response and the later @fig:chapter-two-pipeline are referenced from the first chapter.

#strong[This sentence exercises a complete bold paragraph.]

The original figure appears in @fig:chapter-one-pipeline. <sec:first-paragraph>

Before nested include.

#include "nested/part.typ"

After nested include.

#formula-note[$q$ is an invented scalar.]

+ First synthetic list item.
+ Second synthetic list item.

The following list starts a new sequence.

+ New sequence first item.

#capfig(
  image("../synthetic-project/figures/synthetic-pipeline.svg", width: 70%),
  caption: [First synthetic pipeline],
  label: <fig:chapter-one-pipeline>,
)

#block(breakable: false)[
The equation inside a layout block remains editable.

$
y = a x + b
$ <chapter-one-model>
]

#captab(
  caption: [First synthetic parameters],
  label: <tab:chapter-one-parameters>,
)[
| Symbol | Value |
| --- | --- |
| $a$ | $2$ |
]
