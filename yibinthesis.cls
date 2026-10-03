% !TeX program = xelatex
\NeedsTeXFormat{LaTeX2e}[2020-10-01]
\ProvidesClass{yibinthesis}[2026/07/14 v0.2.2 Unofficial Yibin University undergraduate thesis class]

\RequirePackage{iftex}
\ifXeTeX\else
  \ClassError{yibinthesis}
    {This class requires XeLaTeX}
    {Compile with XeLaTeX, for example: latexmk -xelatex main.tex.}
\fi

\newif\ifyibin@science
\yibin@sciencefalse
\newif\ifyibin@bibliography
\yibin@bibliographytrue
\newif\ifyibin@strictfonts
\yibin@strictfontsfalse
\newif\ifyibin@proposal
\yibin@proposalfalse
\newif\ifyibin@literaturereview
\yibin@literaturereviewfalse
\newif\ifyibin@documenttypeselected
\yibin@documenttypeselectedfalse
\def\yibin@selecteddocumenttype{thesis}

\newcommand{\yibin@selectdocumenttype}[2]{%
  \ifyibin@documenttypeselected
    \ClassError{yibinthesis}
      {Document type options are mutually exclusive:
       '\yibin@selecteddocumenttype' and '#1'}
      {Select exactly one of thesis, proposal, or literature-review.}%
  \else
    \yibin@documenttypeselectedtrue
    \def\yibin@selecteddocumenttype{#1}%
    #2%
  \fi
}

\DeclareOption{thesis}{%
  \yibin@selectdocumenttype{thesis}{%
    \yibin@proposalfalse\yibin@literaturereviewfalse}}
\DeclareOption{proposal}{%
  \yibin@selectdocumenttype{proposal}{%
    \yibin@proposaltrue\yibin@literaturereviewfalse}}
\DeclareOption{literature-review}{%
  \yibin@selectdocumenttype{literature-review}{%
    \yibin@proposalfalse\yibin@literaturereviewtrue}}
\DeclareOption{humanities}{\yibin@sciencefalse}
\DeclareOption{science}{\yibin@sciencetrue}
\DeclareOption{nobibliography}{\yibin@bibliographyfalse}
\DeclareOption{strictfonts}{\yibin@strictfontstrue}
\DeclareOption*{\PassOptionsToClass{\CurrentOption}{ctexbook}}
\ExecuteOptions{humanities}
\ProcessOptions\relax

\LoadClass[UTF8,a4paper,zihao=-4,oneside,openany,fontset=none]{ctexbook}

\RequirePackage{fontspec}
\RequirePackage{xeCJK}
\RequirePackage{xeCJKfntef}
\RequirePackage{expl3}
\RequirePackage{xparse}
\RequirePackage{geometry}
\RequirePackage{setspace}
\RequirePackage{fancyhdr}
\RequirePackage{graphicx}
\RequirePackage{tabularx}
\RequirePackage{array}
\RequirePackage{booktabs}
\RequirePackage{longtable}
\RequirePackage{amsmath}
\RequirePackage{amssymb}
\RequirePackage{caption}
\RequirePackage{chngcntr}
\RequirePackage{titlesec}
\RequirePackage{tocloft}
\RequirePackage{enumitem}
\RequirePackage[absolute,overlay]{textpos}
\ifyibin@bibliography
  \RequirePackage[
    backend=biber,
    style=gb7714-2015,
    sorting=none,
    giveninits=true,
    defernumbers=false
  ]{biblatex}
\fi
\RequirePackage[
  unicode=true,
  bookmarksnumbered=true,
  bookmarksopen=true,
  hypertexnames=false,
  colorlinks=false,
  hidelinks,
  pdfborder={0 0 0}
]{hyperref}
\RequirePackage{bookmark}

% Public table columns use array's m semantics so every paragraph-like cell
% participates in vertical centering.  tabularx's X column is also switched
% from its default p/vtop implementation to m; mixing p/X with m in one row
% makes short cells appear top-aligned even when only those cells request m.
\newcolumntype{L}[1]{>{\setlength{\parindent}{0pt}\raggedright\arraybackslash}m{#1}}
\newcolumntype{C}[1]{>{\setlength{\parindent}{0pt}\centering\arraybackslash}m{#1}}
\newcolumntype{R}[1]{>{\setlength{\parindent}{0pt}\raggedleft\arraybackslash}m{#1}}
\renewcommand{\tabularxcolumn}[1]{>{\setlength{\parindent}{0pt}}m{#1}}

% Page geometry follows the school specification: A4, single-sided,
% 3.0 cm binding side and 2.5 cm on the other three sides.
\geometry{
  a4paper,
  top=2.5cm,
  bottom=2.5cm,
  left=3.0cm,
  right=2.5cm,
  headheight=15pt,
  headsep=10mm,
  footskip=10mm
}

% Prefer the fonts named by the school template.  The default compatibility
% mode falls back to TeX Live fonts; strictfonts turns each missing school font
% into an explicit class error while retaining a fallback so the log is useful.
\newcommand{\yibin@missingfont}[1]{%
  \ifyibin@strictfonts
    \ClassError{yibinthesis}
      {Required font '#1' was not found}
      {Install the school-required font or remove the strictfonts option for a preview build.}%
  \else
    \ClassWarning{yibinthesis}
      {Font '#1' was not found; using a compatible preview font}%
  \fi
}

\IfFontExistsTF{SimSun}{
  \setCJKmainfont{SimSun}[AutoFakeBold=2.5]
  \setCJKmonofont{SimSun}[AutoFakeBold=2.5]
  \setCJKfamilyfont{yibin-song}{SimSun}[AutoFakeBold=2.5]
}{
  \yibin@missingfont{SimSun}
  \setCJKmainfont{FandolSong-Regular}[
    BoldFont=FandolSong-Bold
  ]
  \setCJKmonofont{FandolSong-Regular}[
    BoldFont=FandolSong-Bold
  ]
  \setCJKfamilyfont{yibin-song}{FandolSong-Regular}[
    BoldFont=FandolSong-Bold
  ]
}

\IfFontExistsTF{SimHei}{
  \setCJKsansfont{SimHei}[AutoFakeBold=2.5]
  \setCJKfamilyfont{yibin-hei}{SimHei}[AutoFakeBold=2.5]
}{
  \yibin@missingfont{SimHei}
  \setCJKsansfont{FandolHei-Regular}[
    BoldFont=FandolHei-Bold
  ]
  \setCJKfamilyfont{yibin-hei}{FandolHei-Regular}[
    BoldFont=FandolHei-Bold
  ]
}

\IfFontExistsTF{KaiTi}{
  \setCJKfamilyfont{yibin-kai}{KaiTi}[AutoFakeBold=2.5]
}{
  \yibin@missingfont{KaiTi}
  \setCJKfamilyfont{yibin-kai}{FandolKai-Regular}[AutoFakeBold=2.5]
}

\IfFontExistsTF{Times New Roman}{
  \setmainfont{Times New Roman}
  \newfontfamily\yibinenglishfont{Times New Roman}
}{
  \yibin@missingfont{Times New Roman}
  \setmainfont{TeX Gyre Termes}
  \newfontfamily\yibinenglishfont{TeX Gyre Termes}
}

\DeclareRobustCommand{\yibinsong}{\CJKfamily{yibin-song}}
\DeclareRobustCommand{\yibinhei}{\CJKfamily{yibin-hei}}
\DeclareRobustCommand{\yibinkai}{\CJKfamily{yibin-kai}}

\setstretch{1.5}
\setlength{\parindent}{2em}
\setlength{\parskip}{0pt}
\setcounter{secnumdepth}{3}
\setcounter{tocdepth}{2}
\raggedbottom

\ExplSyntaxOn
\tl_new:N \g_yibin_title_tl
\tl_new:N \g_yibin_english_title_tl
\tl_new:N \g_yibin_author_tl
\tl_new:N \g_yibin_student_id_tl
\tl_new:N \g_yibin_college_tl
\tl_new:N \g_yibin_major_tl
\tl_new:N \g_yibin_grade_tl
\tl_new:N \g_yibin_class_name_tl
\tl_new:N \g_yibin_advisor_tl
\tl_new:N \g_yibin_advisor_title_tl
\tl_new:N \g_yibin_external_advisor_tl
\tl_new:N \g_yibin_external_advisor_title_tl
\tl_new:N \g_yibin_version_tl
\tl_new:N \g_yibin_date_tl
\tl_new:N \g_yibin_declassify_year_tl
\tl_new:N \g_yibin_logo_tl
\tl_new:N \g_yibin_author_signature_tl
\tl_new:N \g_yibin_advisor_signature_tl
\tl_new:N \g_yibin_signature_background_tl
\tl_new:N \g_yibin_template_year_tl
\tl_new:N \g_yibin_expected_template_year_tl
\bool_new:N \g_yibin_confidential_bool

% The opening report is the form extracted from the 2022 work-form.  Thesis
% and literature-review documents use the 2024 thesis typography baseline.
\ifyibin@proposal
  \tl_gset:Nn \g_yibin_expected_template_year_tl {2022}
\else
  \tl_gset:Nn \g_yibin_expected_template_year_tl {2024}
\fi
\tl_gset:Nx \g_yibin_template_year_tl
  {\tl_use:N \g_yibin_expected_template_year_tl}

\keys_define:nn { yibin / metadata } {
  title                    .tl_gset:N = \g_yibin_title_tl,
  english-title            .tl_gset:N = \g_yibin_english_title_tl,
  author                   .tl_gset:N = \g_yibin_author_tl,
  student-id               .tl_gset:N = \g_yibin_student_id_tl,
  college                  .tl_gset:N = \g_yibin_college_tl,
  major                    .tl_gset:N = \g_yibin_major_tl,
  grade                    .tl_gset:N = \g_yibin_grade_tl,
  class-name               .tl_gset:N = \g_yibin_class_name_tl,
  advisor                  .tl_gset:N = \g_yibin_advisor_tl,
  advisor-title            .tl_gset:N = \g_yibin_advisor_title_tl,
  external-advisor         .tl_gset:N = \g_yibin_external_advisor_tl,
  external-advisor-title   .tl_gset:N = \g_yibin_external_advisor_title_tl,
  version                  .tl_gset:N = \g_yibin_version_tl,
  date                     .tl_gset:N = \g_yibin_date_tl,
  template-year            .code:n = {
    \tl_set:Nx \l_tmpa_tl {\tl_trim_spaces:n {#1}}
    \tl_if_eq:NNTF \l_tmpa_tl \g_yibin_expected_template_year_tl
      {\tl_gset:Nx \g_yibin_template_year_tl {\tl_use:N \l_tmpa_tl}}
      {
        \ClassError{yibinthesis}
          {Invalid template-year '#1' for document type
           '\yibindocumenttype'}
          {Document type '\yibindocumenttype' only accepts
           template-year=\tl_use:N \g_yibin_expected_template_year_tl.}
      }
  },
  declassify-year          .tl_gset:N = \g_yibin_declassify_year_tl,
  logo                     .code:n = {
    \str_if_eq:nnTF {#1} {builtin}
      {\tl_gset:Nn \g_yibin_logo_tl {assets/yibin-university-logo.png}}
      {\tl_gset:Nn \g_yibin_logo_tl {#1}}
  },
  author-signature         .tl_gset:N = \g_yibin_author_signature_tl,
  advisor-signature        .tl_gset:N = \g_yibin_advisor_signature_tl,
  signature-background     .choice:,
  signature-background / preserve .code:n =
    \tl_gset:Nn \g_yibin_signature_background_tl {preserve},
  signature-background / whiten .code:n =
    \tl_gset:Nn \g_yibin_signature_background_tl {whiten},
  signature-background     .initial:n = preserve,
  secrecy                  .choice:,
  secrecy / public         .code:n = \bool_gset_false:N \g_yibin_confidential_bool,
  secrecy / confidential   .code:n = \bool_gset_true:N \g_yibin_confidential_bool,
  secrecy                  .initial:n = public,
  unknown                  .code:n = {
    \ClassWarning{yibinthesis}{Unknown metadata key '\l_keys_key_str'}
  }
}

\keys_set:nn { yibin / metadata } {
  title                   = {（填写中文题目）},
  english-title           = {Title in English},
  author                  = {（填写姓名）},
  student-id              = {（填写学号）},
  college                 = {（填写学院）},
  major                   = {（填写专业）},
  grade                   = {（填写年级）},
  class-name              = {},
  advisor                 = {（填写校内导师）},
  advisor-title           = {（填写职称）},
  external-advisor        = {},
  external-advisor-title  = {},
  version                 = {},
  date                    = {},
  declassify-year         = {},
  logo                    = {builtin},
  author-signature        = {},
  advisor-signature       = {},
  signature-background    = preserve,
  secrecy                 = public
}

\NewDocumentCommand{\yibinsetup}{m}{
  \keys_set:nn { yibin / metadata } {#1}
}

\NewExpandableDocumentCommand{\yibintitle}{}{\tl_use:N \g_yibin_title_tl}
\NewExpandableDocumentCommand{\yibinenglishtitle}{}{\tl_use:N \g_yibin_english_title_tl}
\NewExpandableDocumentCommand{\yibinauthor}{}{\tl_use:N \g_yibin_author_tl}
\NewExpandableDocumentCommand{\yibinstudentid}{}{\tl_use:N \g_yibin_student_id_tl}
\NewExpandableDocumentCommand{\yibincollege}{}{\tl_use:N \g_yibin_college_tl}
\NewExpandableDocumentCommand{\yibinmajor}{}{\tl_use:N \g_yibin_major_tl}
\NewExpandableDocumentCommand{\yibingrade}{}{\tl_use:N \g_yibin_grade_tl}
\NewExpandableDocumentCommand{\yibinclassname}{}{\tl_use:N \g_yibin_class_name_tl}
\NewDocumentCommand{\yibingradeclass}{}{
  \tl_if_empty:NF \g_yibin_grade_tl {
    \tl_use:N \g_yibin_grade_tl
    \tl_if_in:NnF \g_yibin_grade_tl {级} {级}
  }
  \tl_use:N \g_yibin_class_name_tl
}
\NewExpandableDocumentCommand{\yibinadvisor}{}{\tl_use:N \g_yibin_advisor_tl}
\NewExpandableDocumentCommand{\yibinadvisortitle}{}{\tl_use:N \g_yibin_advisor_title_tl}
\NewExpandableDocumentCommand{\yibinexternaladvisor}{}{\tl_use:N \g_yibin_external_advisor_tl}
\NewExpandableDocumentCommand{\yibinexternaladvisortitle}{}{\tl_use:N \g_yibin_external_advisor_title_tl}
\NewExpandableDocumentCommand{\yibinversion}{}{\tl_use:N \g_yibin_version_tl}
\NewExpandableDocumentCommand{\yibindate}{}{\tl_use:N \g_yibin_date_tl}
\NewDocumentCommand{\yibinifversion}{mm}{
  \tl_if_empty:NTF \g_yibin_version_tl {#2} {#1}
}
\NewDocumentCommand{\yibinifdate}{mm}{
  \tl_if_empty:NTF \g_yibin_date_tl {#2} {#1}
}
\NewExpandableDocumentCommand{\yibintemplateyear}{}{\tl_use:N \g_yibin_template_year_tl}
\NewExpandableDocumentCommand{\yibindeclassifyyear}{}{\tl_use:N \g_yibin_declassify_year_tl}
\NewExpandableDocumentCommand{\yibinlogo}{}{\tl_use:N \g_yibin_logo_tl}
\NewExpandableDocumentCommand{\yibinauthorsignature}{}{\tl_use:N \g_yibin_author_signature_tl}
\NewExpandableDocumentCommand{\yibinadvisorsignature}{}{\tl_use:N \g_yibin_advisor_signature_tl}
\NewExpandableDocumentCommand{\yibinsignaturebackground}{}{\tl_use:N \g_yibin_signature_background_tl}

\NewDocumentCommand{\yibinsecretbox}{}{
  \bool_if:NTF \g_yibin_confidential_bool {$\boxtimes$} {$\square$}
}
\NewDocumentCommand{\yibinpublicbox}{}{
  \bool_if:NTF \g_yibin_confidential_bool {$\square$} {$\boxtimes$}
}
\ExplSyntaxOff

\newcommand{\yibindocumenttype}{\yibin@selecteddocumenttype}

\newcommand{\yibin@documentsubject}{%
  \ifyibin@proposal
    宜宾学院本科毕业论文（设计）开题报告%
  \else
    \ifyibin@literaturereview
      宜宾学院本科毕业论文（设计）文献综述%
    \else
      宜宾学院本科生毕业论文（设计）%
    \fi
  \fi
}

\AtBeginDocument{
  % The build tool may provide derived (for example, whitened) signature
  % images here.  Source metadata and original image files remain untouched.
  \InputIfFileExists{yibinthesis-build-overrides.tex}{}{}
  \hypersetup{
    pdftitle={\yibintitle},
    pdfauthor={\yibinauthor},
    pdfsubject={\yibin@documentsubject / \yibinenglishtitle},
    pdfcreator={XeLaTeX with yibinthesis class}
  }
}

% Heading hierarchy. The class option controls numbering only; the font
% hierarchy follows the same school-level table for both discipline groups.
\ifyibin@science
  \renewcommand{\thechapter}{\arabic{chapter}}
  \renewcommand{\thesection}{\thechapter.\arabic{section}}
  \renewcommand{\thesubsection}{\thesection.\arabic{subsection}}
  \renewcommand{\thesubsubsection}{\thesubsection.\arabic{subsubsection}}
  \ctexset{
    chapter = {
      name = {,},
      number = \thechapter,
      aftername = {\quad}
    }
  }
\else
  \renewcommand{\thechapter}{\chinese{chapter}、}
  \renewcommand{\thesection}{（\chinese{section}）}
  \renewcommand{\thesubsection}{\arabic{subsection}.}
  \renewcommand{\thesubsubsection}{（\arabic{subsubsection}）}
  \ctexset{
    chapter = {
      name = {,},
      number = \thechapter,
      aftername = {\quad}
    }
  }
\fi

\ifyibin@science
  \titleformat{\chapter}[block]
    {\centering\yibinhei\bfseries\zihao{3}}
    {\thechapter}{0.5em}{}
  \titlespacing{\chapter}{0pt}{0pt}{24pt}
\else
  % Humanities level-one headings remain left aligned with a two-character
  % indent, but each numbered chapter starts on a new page after the opening.
  \titleclass{\chapter}{top}
  \titleformat{\chapter}[block]
    {\yibinhei\bfseries\zihao{3}}
    {\thechapter}{0.5em}{}
  \titlespacing{\chapter}{2em}{18pt}{12pt}
\fi

% Abstracts, the table of contents, introduction/conclusion and back matter
% remain centered three-size headings in both discipline profiles.
\titleformat{name=\chapter,numberless}[block]
  {\centering\yibinhei\bfseries\zihao{3}}
  {}{0pt}{}
\titlespacing{name=\chapter,numberless}{0pt}{0pt}{24pt}

\titleformat{\section}[block]
  {\yibinkai\bfseries\zihao{-3}}
  {\thesection}{0.5em}{}
% Use the unstarred form at every level.  In titlesec the starred form
% suppresses the indentation of the paragraph immediately following a title,
% which conflicts with the required two-character body first-line indent.
\titlespacing{\section}{2em}{18pt}{12pt}

\titleformat{\subsection}[block]
  {\yibinsong\bfseries\zihao{4}}
  {\thesubsection}{0.5em}{}
\titlespacing{\subsection}{2em}{15pt}{9pt}

\titleformat{\subsubsection}[block]
  {\yibinsong\bfseries\zihao{4}}
  {\thesubsubsection}{0.5em}{}
\titlespacing{\subsubsection}{2em}{12pt}{6pt}

\renewcommand{\contentsname}{目录}
\renewcommand{\listfigurename}{图目录}
\renewcommand{\listtablename}{表目录}
\renewcommand{\cfttoctitlefont}{\hfill\yibinhei\bfseries\zihao{3}}
% Keep material after the trailing fill so TeX does not discard it at the
% paragraph boundary and accidentally right-align the title.
\renewcommand{\cftaftertoctitle}{\hfill\mbox{}}
\setlength{\cftbeforetoctitleskip}{0pt}
\setlength{\cftaftertoctitleskip}{24pt}
\renewcommand{\cftloftitlefont}{\hfill\yibinhei\bfseries\zihao{3}}
\renewcommand{\cftafterloftitle}{\hfill\mbox{}}
\setlength{\cftbeforeloftitleskip}{0pt}
\setlength{\cftafterloftitleskip}{24pt}
\renewcommand{\cftlottitlefont}{\hfill\yibinhei\bfseries\zihao{3}}
\renewcommand{\cftafterlottitle}{\hfill\mbox{}}
\setlength{\cftbeforelottitleskip}{0pt}
\setlength{\cftafterlottitleskip}{24pt}
\renewcommand{\cftchapfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftchappagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsecfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsecpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsubsecfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsubsecpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftchapleader}{\cftdotfill{\cftdotsep}}
\renewcommand{\cftfigfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftfigpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftfigleader}{\cftdotfill{\cftdotsep}}
\renewcommand{\cfttabfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cfttabpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cfttableader}{\cftdotfill{\cftdotsep}}
\setlength{\cftbeforechapskip}{0.25\baselineskip}
% The official Word template uses absolute left indents of 0/21/42 pt for
% levels one to three.  At the 12 pt TOC size these map to 0/1.75/3.5 em.
% Set them explicitly instead of inheriting tocloft's larger defaults.
\setlength{\cftchapindent}{0em}
\setlength{\cftsecindent}{1.75em}
\setlength{\cftsubsecindent}{3.5em}
\ifyibin@science
  \setlength{\cftchapnumwidth}{2.5em}
  \setlength{\cftsecnumwidth}{3.5em}
  \setlength{\cftsubsecnumwidth}{4.5em}
\else
  % Number boxes follow the visible widths of “一、”, “（一）” and “1.”.
  % Oversized boxes make the title text look like an additional TOC indent.
  \setlength{\cftchapnumwidth}{2em}
  \setlength{\cftsecnumwidth}{3em}
  \setlength{\cftsubsecnumwidth}{1.5em}
\fi

% The 2024 Word template gives the table of contents its own unnumbered
% section.  Keep its pages visually unnumbered while preserving the Roman
% counter used by the preceding abstracts.
\newcommand{\yibin@tableofcontents}{%
  \clearpage
  % Keep these assignments active until the TOC page has actually shipped.
  % A local group can end before fancyhdr expands its footer in some engines.
  \let\yibin@savedpsplain\ps@plain
  \let\ps@plain\ps@empty
  \pagestyle{empty}
  \yibin@originaltableofcontents
  \thispagestyle{empty}
  \clearpage
  \let\ps@plain\yibin@savedpsplain
  \pagestyle{yibinfront}
}

% hyperref/bookmark finish patching \tableofcontents at begin-document time.
% Install the visible-page-style wrapper afterwards so it cannot be replaced.
\AddToHook{begindocument/end}{%
  \let\yibin@originaltableofcontents\tableofcontents
  \let\tableofcontents\yibin@tableofcontents
}

% Front matter uses upper-case Roman numerals with no header. Main matter
% restarts at Arabic 1 and carries the thesis title in the header.
\fancypagestyle{yibinfront}{
  \fancyhf{}
  \fancyfoot[C]{\yibinenglishfont\zihao{-5}\thepage}
  \renewcommand{\headrulewidth}{0pt}
  \renewcommand{\footrulewidth}{0pt}
}

\fancypagestyle{yibinmain}{
  \fancyhf{}
  \fancyhead[C]{\yibin@headertitle}
  \fancyfoot[C]{\yibinenglishfont\zihao{-5}\thepage}
  \renewcommand{\headrulewidth}{0.4pt}
  \renewcommand{\footrulewidth}{0pt}
}

\newcommand{\yibin@plainfront}{%
  \fancypagestyle{plain}{
    \fancyhf{}
    \fancyfoot[C]{\yibinenglishfont\zihao{-5}\thepage}
    \renewcommand{\headrulewidth}{0pt}
    \renewcommand{\footrulewidth}{0pt}
  }
}

\newcommand{\yibin@plainmain}{%
  \fancypagestyle{plain}{
    \fancyhf{}
    \fancyhead[C]{\yibin@headertitle}
    \fancyfoot[C]{\yibinenglishfont\zihao{-5}\thepage}
    \renewcommand{\headrulewidth}{0.4pt}
    \renewcommand{\footrulewidth}{0pt}
  }
}

\newsavebox{\yibin@headerbox}
\newcommand{\yibin@headertitle}{%
  \sbox{\yibin@headerbox}{\yibinsong\zihao{-5}\yibintitle}%
  \ifdim\wd\yibin@headerbox>\headwidth
    \resizebox{\headwidth}{!}{\usebox{\yibin@headerbox}}%
  \else
    \usebox{\yibin@headerbox}%
  \fi
}

\renewcommand{\frontmatter}{
  \clearpage
  \@mainmatterfalse
  \pagenumbering{Roman}
  \pagestyle{yibinfront}
  \yibin@plainfront
}

\renewcommand{\mainmatter}{
  \clearpage
  \@mainmattertrue
  \pagenumbering{arabic}
  \pagestyle{yibinmain}
  \yibin@plainmain
}

\renewcommand{\backmatter}{
  \clearpage
  \@mainmatterfalse
  \pagestyle{yibinmain}
  \yibin@plainmain
}

\newsavebox{\yibin@covertitlebox}
\newcommand{\yibin@covertitle}{%
  \begingroup
    \yibinhei\bfseries\fontsize{18bp}{26bp}\selectfont
    \sbox{\yibin@covertitlebox}{\strut\yibintitle\strut}%
    \ifdim\wd\yibin@covertitlebox>15cm
      % Keep unusually long real titles usable.  The official one-line
      % placeholder remains a naturally underlined run below.  A slightly
      % narrower inner measure avoids a one-character final line.
      \parbox{13.2cm}{\centering\CJKunderline{\yibintitle}\par}%
    \else
      \underline{\usebox{\yibin@covertitlebox}}%
    \fi
  \endgroup
}

\newcommand{\yibin@coverlabel}[2]{%
  \makebox[#1][l]{\yibinhei\bfseries\fontsize{18bp}{22bp}\selectfont #2}%
}

\newsavebox{\yibin@covervaluebox}
\newcommand{\yibin@covervalue}[2]{%
  \begingroup
    \yibinsong\bfseries\fontsize{16bp}{20bp}\selectfont
    \sbox{\yibin@covervaluebox}{\strut #2\strut}%
    \underline{\makebox[#1][c]{%
      \ifdim\wd\yibin@covervaluebox>#1
        \resizebox{\dimexpr#1-0.2cm\relax}{!}{\usebox{\yibin@covervaluebox}}%
      \else
        \usebox{\yibin@covervaluebox}%
      \fi}}%
  \endgroup
}

\newcommand{\yibin@coverfield}[4]{%
  \noindent\yibin@coverlabel{#2}{#1}\yibin@covervalue{#3}{#4}%
}

\newcommand{\yibin@coverdualfield}[8]{%
  \noindent
  \yibin@coverlabel{#2}{#1}\yibin@covervalue{#3}{#4}%
  \yibin@coverlabel{#6}{#5}\yibin@covervalue{#7}{#8}%
}

\NewDocumentCommand{\makeyibincover}{}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    \setlength{\parindent}{0pt}
    % Fixed cover elements are positioned from the physical page edge so
    % their coordinates are independent of paragraph spacing and title wraps.
    \begin{textblock*}{13.15cm}(3.925cm,2.606cm)
      \IfFileExists{\yibinlogo}{%
        \includegraphics[width=13.15cm,height=3.65cm]{\yibinlogo}%
      }{%
        \parbox[c][3.65cm][c]{13.15cm}{\centering
          {\yibinhei\bfseries\fontsize{34bp}{42bp}\selectfont 宜宾学院\par}
          {\yibinenglishfont\bfseries\fontsize{18bp}{22bp}\selectfont YIBIN UNIVERSITY\par}}
      }
    \end{textblock*}

    \begin{textblock*}{21cm}(0cm,7.064cm)
      \centering\yibinhei\bfseries\fontsize{28bp}{34bp}\selectfont
      本科生毕业论文（设计）
    \end{textblock*}

    \begin{textblock*}{15cm}(3cm,10.160cm)
      \centering\yibin@covertitle
    \end{textblock*}

    \yibinifversion{%
      \begin{textblock*}{21cm}(0cm,12.55cm)
        \centering\yibinhei\zihao{4}\yibinversion
      \end{textblock*}%
    }{}

    \begin{textblock*}{15.5cm}(3cm,13.447cm)
      \yibin@coverfield{学院（部）}{3.5cm}{11.6cm}{\yibincollege}
    \end{textblock*}
    \begin{textblock*}{15.5cm}(3cm,15.109cm)
      \yibin@coverfield{专\hspace{2em}业}{3.5cm}{11.6cm}{\yibinmajor}
    \end{textblock*}
    \begin{textblock*}{15.5cm}(3cm,16.733cm)
      \yibin@coverfield{学生姓名}{3.2cm}{11.9cm}{\yibinauthor}
    \end{textblock*}
    \begin{textblock*}{15.5cm}(3cm,18.357cm)
      \yibin@coverdualfield{学\hspace{2em}号}{3.2cm}{5.6cm}{\yibinstudentid}%
        {年级}{1.6cm}{4.5cm}{\yibingradeclass}
    \end{textblock*}
    \begin{textblock*}{15.5cm}(3cm,20.019cm)
      \yibin@coverdualfield{指导教师（校内）}{5.1cm}{5.1cm}{\yibinadvisor}%
        {职称}{1.55cm}{3.1cm}{\yibinadvisortitle}
    \end{textblock*}
    \begin{textblock*}{15.5cm}(3cm,21.643cm)
      \yibin@coverdualfield{指导教师（校外）}{5.1cm}{5.1cm}{\yibinexternaladvisor}%
        {职称}{1.55cm}{3.1cm}{\yibinexternaladvisortitle}
    \end{textblock*}

    \yibinifdate{%
      \begin{textblock*}{21cm}(0cm,25.55cm)
        \centering\yibinsong\bfseries\zihao{-3}\yibindate
      \end{textblock*}%
    }{}
    % textblock* content does not enter the ordinary page list.  Keep one
    % real page object so the fixed cover blocks ship before the declaration.
    \null
  \endgroup
  \clearpage
}

\newcommand{\yibin@signatureimage}[2]{%
  \begingroup
    \edef\yibin@signaturepath{#1}%
    \ifx\yibin@signaturepath\@empty\else
      \IfFileExists{\yibin@signaturepath}{%
        \raisebox{-0.12cm}[0.92cm][0.04cm]{%
          \includegraphics[width=#2,height=1.08cm,keepaspectratio]{\yibin@signaturepath}}%
      }{}%
    \fi
  \endgroup
}

\newcommand{\yibin@signatureline}[2]{%
  #1\hspace{0.5em}%
  \underline{\makebox[3.0cm][c]{\yibin@signatureimage{#2}{2.8cm}}}%
  \ignorespaces
}

% The copyright page uses two fixed columns in the official Word template.
% Keep this variant separate from the originality-page signature line: the
% official copyright lines are shorter (about 2.5 cm) and must not drift when
% the left-hand label changes width.
\newcommand{\yibin@copyrightsignatureline}[2]{%
  #1\hspace{0.05em}%
  \underline{\makebox[2.5cm][c]{\yibin@signatureimage{#2}{2.35cm}}}%
  \ignorespaces
}

\newcommand{\makeyibinoriginality}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    \begin{textblock*}{15.5cm}(3cm,2.75cm)
      \centering\yibinhei\bfseries\fontsize{16bp}{24bp}\selectfont 原创性声明
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,5.05cm)
      \yibinsong\fontsize{12bp}{14.4bp}\selectfont
      \setlength{\parindent}{24bp}\setstretch{2.0}
      本人呈交的学位论文（设计）《\yibintitle》，是在导师的指导下，独立进行研究取得的成果。
      除文中已经注明引用的内容外，本论文（设计）不包括其他个人或集体已经发表或撰写过的作品成果。
      对本文（设计）做出贡献的个人和集体，均已在文中以明确方式标明。
      本人完全意识到本声明的法律后果，因本声明而产生的法律后果由本人承担。
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,12.88cm)
      \yibinsong\fontsize{12bp}{24bp}\selectfont
      \noindent\hspace*{24bp}\yibin@signatureline{学位论文作者：}{\yibinauthorsignature}
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,14.50cm)
      \yibinsong\fontsize{12bp}{24bp}\selectfont
      \noindent\hspace*{24bp}日期：\hspace{2em}年\hspace{2em}月\hspace{2em}日
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,21.87cm)
      \noindent
      {\yibinsong\bfseries\fontsize{16bp}{24bp}\selectfont 附：}%
      {\yibinsong\bfseries\fontsize{14bp}{21bp}\selectfont
        《普通高等学校学生管理规定》（中华人民共和国教育部令第41号）}
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,23.00cm)
      \yibinsong\bfseries\fontsize{14bp}{21bp}\selectfont\setstretch{1.5}
      \noindent\hspace*{35bp}第五十二条~学生有下列情形之一，学校可以给予开除学籍处分：
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,24.10cm)
      \yibinsong\bfseries\fontsize{14bp}{21bp}\selectfont\setstretch{1.5}
      \setlength{\leftskip}{7.15bp}\noindent
      （五）学位论文、公开发表的研究成果存在抄袭、篡改、伪造等学术不端行为，情节严重的，或者代写论文、买卖论文的；
    \end{textblock*}
    \null
  \endgroup
  \clearpage
}

\newcommand{\makeyibincopyright}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    % The official page is laid out as fixed vertical slots (title 2.75 cm,
    % body 5.05 cm, choices 10.15--11.82 cm, signatures 15.10/16.75 cm).
    % Absolute text blocks avoid the cumulative paragraph-spacing drift that
    % otherwise moves the signatures toward the bottom of the page.
    \begin{textblock*}{15.5cm}(3cm,2.75cm)
      \centering\yibinhei\bfseries\fontsize{16bp}{24bp}\selectfont
      学位论文（设计）版权使用授权书
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,5.05cm)
      \yibinsong\fontsize{12bp}{24bp}\selectfont
      % Explicit baseline matches the official 24-pt Word line spacing after
      % the CJK font metrics are applied by XeTeX.
      \setlength{\baselineskip}{31.2pt}
      \setlength{\parindent}{24bp}
      本学位论文（设计）作者完全了解学校有关保留、使用学位论文（设计）的规定，
      同意学校保留并向国家有关部门或机构送交论文（设计）的复印件和电子版，
      允许论文（设计）被查阅和借阅。本人授权宜宾学院将本学位论文（设计）的全部或部分内容
      编入有关数据库进行检索，可以采用影印、缩印或扫描等复制手段保存和汇编。
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,10.15cm)
      \noindent\hspace*{24bp}\yibinsong\fontsize{12bp}{18bp}\selectfont
      本学位论文（设计）属于\underline{（请在以下相应方框内打“√”）作品}
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,10.98cm)
      \noindent\hspace*{24bp}\yibinsong\fontsize{12bp}{18bp}\selectfont
      保\hspace{1em}密\yibinsecretbox，在\underline{\makebox[2.0cm][c]{\yibindeclassifyyear}}年解密后适用本授权书。
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,11.82cm)
      \noindent\hspace*{24bp}\yibinsong\fontsize{12bp}{18bp}\selectfont
      不保密\yibinpublicbox。
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,15.10cm)
      \noindent\hspace*{24bp}\yibinsong\fontsize{12bp}{18bp}\selectfont
      \makebox[6.38cm][l]{\yibin@copyrightsignatureline{作者（签名）：}{\yibinauthorsignature}}%
      \makebox[6.38cm][l]{\yibin@copyrightsignatureline{指导教师（签名）：}{\yibinadvisorsignature}}%
    \end{textblock*}

    \begin{textblock*}{15.5cm}(3cm,16.75cm)
      \noindent\hspace*{24bp}\yibinsong\fontsize{12bp}{18bp}\selectfont
      \makebox[6.38cm][l]{日期：\hspace{0.55em}年\hspace{1.45em}月\hspace{1.45em}日}%
      \makebox[6.38cm][l]{\hspace*{11pt}日期：\hspace*{12pt}\hspace{0.55em}年\hspace{1.45em}月\hspace{1.45em}日}%
    \end{textblock*}
    \null
  \endgroup
  \clearpage
}

\NewDocumentCommand{\makeyibindeclarations}{}{%
  \makeyibinoriginality
  \makeyibincopyright
}

\NewDocumentCommand{\yibinunnumberedchapter}{m}{%
  \clearpage
  \setcounter{section}{0}
  \setcounter{subsection}{0}
  \setcounter{subsubsection}{0}
  \chapter*{#1}
  \phantomsection
  \addcontentsline{toc}{chapter}{#1}
  \markboth{#1}{#1}
}

% Every appendix is a top-level unnumbered chapter.  This keeps the official
% black 16 pt centered title and guarantees a fresh page for A, B, C, ... .
\newcounter{yibinappendix}
\NewDocumentEnvironment{yibinappendices}{}{%
  \setcounter{yibinappendix}{0}%
}{%
  \par
}
\NewDocumentCommand{\yibinappendix}{m}{%
  \refstepcounter{yibinappendix}%
  \clearpage
  \setcounter{section}{0}%
  \setcounter{subsection}{0}%
  \setcounter{subsubsection}{0}%
  \ifyibin@science
    \renewcommand{\thesection}{\Alph{yibinappendix}.\arabic{section}}%
  \else
    \renewcommand{\thesection}{（\chinese{section}）}%
  \fi
  \chapter*{附录\Alph{yibinappendix}\quad #1}%
  \phantomsection
  \addcontentsline{toc}{chapter}{附录\Alph{yibinappendix}\quad #1}%
  \markboth{附录\Alph{yibinappendix}\quad #1}{附录\Alph{yibinappendix}\quad #1}%
}
\NewDocumentCommand{\yibinappendixsection}{m}{\section{#1}}

% The current humanities sample keeps 绪论 unnumbered, whereas the science
% hierarchy requires 绪论 to participate in chapter numbering.  The school
% writing specification requires 结论 to be a separate unnumbered chapter in
% every discipline profile; \yibinunnumberedchapter also keeps it in the TOC.
\NewDocumentCommand{\yibinopeningchapter}{O{绪论}}{%
  \ifyibin@science\chapter{#1}\else\yibinunnumberedchapter{#1}\fi
}
\NewDocumentCommand{\yibinclosingchapter}{O{结论}}{%
  \yibinunnumberedchapter{#1}
}

\NewDocumentEnvironment{cnabstract}{}{%
  \yibinunnumberedchapter{摘要}
  \yibinsong\zihao{-4}
}{%
  \par
}

\NewDocumentCommand{\cnkeywords}{m}{%
  \par\vspace{\baselineskip}
  \noindent{\yibinhei\bfseries 关键词：}{\yibinsong #1}\par
}

\NewDocumentEnvironment{enabstract}{}{%
  \yibinunnumberedchapter{Abstract}
  \begingroup
  \yibinenglishfont\zihao{-4}
  \setlength{\parindent}{0pt}
}{%
  \par
  \endgroup
}

\NewDocumentCommand{\enkeywords}{m}{%
  \par\vspace{\baselineskip}
  \noindent{\yibinenglishfont\bfseries Keywords: }
  {\yibinenglishfont #1}\par
}

% Caption labels are five-size Song type and centered. Put \caption before
% a table body and after a figure body, as demonstrated in the sample chapter.
\DeclareCaptionFont{yibinfive}{\yibinsong\zihao{5}}
\captionsetup{
  font=yibinfive,
  labelfont=yibinfive,
  labelsep=space,
  justification=centering,
  singlelinecheck=true
}
\captionsetup[table]{position=top}
\captionsetup[figure]{position=bottom}

% Thesis float-only pages read more naturally from the top.  LaTeX's default
% infinite stretch above and between floats can otherwise place a single table
% in the vertical centre or leave a large gap between a figure and a table.
\setlength{\@fptop}{0pt}
\setlength{\@fpsep}{12pt plus 2pt minus 2pt}
\setlength{\@fpbot}{0pt plus 1fil}

\ifyibin@science
  \counterwithin{figure}{chapter}
  \counterwithin{table}{chapter}
\else
  \counterwithout{figure}{chapter}
  \counterwithout{table}{chapter}
\fi
\numberwithin{equation}{chapter}
\renewcommand{\theequation}{\arabic{chapter}.\arabic{equation}}

\newcounter{yibinnote}
\ExplSyntaxOn
\NewDocumentCommand{\yibinnote}{m}{
  \refstepcounter{yibinnote}
  \textsuperscript{\textcircled{\scriptsize\arabic{yibinnote}}}
  \cs_gset:cpx
    { yibin_note_\int_eval:n { \value{yibinnote} } : }
    { \exp_not:n {#1} }
}

\NewDocumentCommand{\printyibinnotes}{}{
  \int_compare:nNnT { \value{yibinnote} } > { 0 } {
    \yibinunnumberedchapter{注释}
    \begingroup
      \yibinsong\zihao{5}
      \int_step_inline:nn { \value{yibinnote} } {
        \par\noindent
        \textcircled{\scriptsize ##1}\quad
        \use:c { yibin_note_##1: }
        \par
      }
    \endgroup
  }
}
\ExplSyntaxOff

\ifyibin@bibliography
  \NewDocumentCommand{\yibincite}{O{}m}{%
    \textsuperscript{\cite[#1]{#2}}%
  }

  \defbibheading{yibinbibliography}[\bibname]{%
    \yibinunnumberedchapter{#1}
  }
  \AtBeginBibliography{\yibinsong\zihao{5}\setstretch{1.5}}
  \NewDocumentCommand{\printyibinbibliography}{}{%
    \printbibliography[heading=yibinbibliography,title={参考文献}]
  }
\else
  \NewDocumentCommand{\addbibresource}{O{}m}{%
    \ClassWarning{yibinthesis}{Bibliography resource '#2' ignored because bibliography is disabled}%
  }
  \NewDocumentCommand{\yibincite}{O{}m}{%
    \ClassWarning{yibinthesis}{Citation '#2' ignored because bibliography is disabled}%
  }
  \NewDocumentCommand{\printyibinbibliography}{}{}
\fi

\setlist{nosep,leftmargin=2em}

% Profile modules share logo and title primitives from this small package.
\RequirePackage{yibinthesis-common}

% Document-specific layouts live in separate modules so the core class keeps
% options, metadata and shared thesis typography only.
\ifyibin@proposal
  \RequirePackage{yibinthesis-proposal}
\fi
\ifyibin@literaturereview
  \RequirePackage{yibinthesis-literature-review}
\fi

\endinput
