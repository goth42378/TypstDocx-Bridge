// Fictional full-thesis regression fixture; no real thesis material.
#show: nwpu-thesis.with(
  graduate: true,
  degree: "doctor",
  track: "academic",
  anonymous: false,
  info: (
    title: ("虚构信号系统", "方法示例"),
    title-en: "A Fictional Signal Processing Study",
    class-no: "T123",
    student-id: "SYNTHETIC",
    author: "示例作者",
    author-en: "Example Author",
    supervisor: ("示例导师", "教授"),
    supervisor-en: "Example Supervisor",
    department: "示例学院",
    major: "示例专业",
    major-en: "Fictional Engineering",
    submit-date: (year: 2026, month: 10),
  ),
  abstract: (
    content: [这是一段虚构的中文摘要，用于验证摘要正文、关键词与页码。],
    keywords: ("虚构信号", "转换测试"),
  ),
  abstract-en: (
    content: [This fictional abstract checks editable text and page numbering.],
    keywords: ("fictional signal", "conversion test"),
  ),
)
#include "chapter-one.typ"
#include "chapter-two.typ"
