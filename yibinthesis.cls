% !TeX program = xelatex
\NeedsTeXFormat{LaTeX2e}[2020-10-01]
\ProvidesClass{yibinthesis}[2026/07/14 v0.2.1 Unofficial Yibin University undergraduate thesis class]

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
\tl_new:N \g_yibin_advisor_tl
\tl_new:N \g_yibin_advisor_title_tl
\tl_new:N \g_yibin_external_advisor_tl
\tl_new:N \g_yibin_external_advisor_title_tl
\tl_new:N \g_yibin_date_tl
\tl_new:N \g_yibin_declassify_year_tl
\tl_new:N \g_yibin_logo_tl
\bool_new:N \g_yibin_confidential_bool

\keys_define:nn { yibin / metadata } {
  title                    .tl_gset:N = \g_yibin_title_tl,
  english-title            .tl_gset:N = \g_yibin_english_title_tl,
  author                   .tl_gset:N = \g_yibin_author_tl,
  student-id               .tl_gset:N = \g_yibin_student_id_tl,
  college                  .tl_gset:N = \g_yibin_college_tl,
  major                    .tl_gset:N = \g_yibin_major_tl,
  grade                    .tl_gset:N = \g_yibin_grade_tl,
  advisor                  .tl_gset:N = \g_yibin_advisor_tl,
  advisor-title            .tl_gset:N = \g_yibin_advisor_title_tl,
  external-advisor         .tl_gset:N = \g_yibin_external_advisor_tl,
  external-advisor-title   .tl_gset:N = \g_yibin_external_advisor_title_tl,
  date                     .tl_gset:N = \g_yibin_date_tl,
  declassify-year          .tl_gset:N = \g_yibin_declassify_year_tl,
  logo                     .tl_gset:N = \g_yibin_logo_tl,
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
  advisor                 = {（填写校内导师）},
  advisor-title           = {（填写职称）},
  external-advisor        = {},
  external-advisor-title  = {},
  date                    = {20XX 年 X 月 X 日},
  declassify-year         = {},
  logo                    = {assets/yibin-university-logo.png},
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
\NewExpandableDocumentCommand{\yibinadvisor}{}{\tl_use:N \g_yibin_advisor_tl}
\NewExpandableDocumentCommand{\yibinadvisortitle}{}{\tl_use:N \g_yibin_advisor_title_tl}
\NewExpandableDocumentCommand{\yibinexternaladvisor}{}{\tl_use:N \g_yibin_external_advisor_tl}
\NewExpandableDocumentCommand{\yibinexternaladvisortitle}{}{\tl_use:N \g_yibin_external_advisor_title_tl}
\NewExpandableDocumentCommand{\yibindate}{}{\tl_use:N \g_yibin_date_tl}
\NewExpandableDocumentCommand{\yibindeclassifyyear}{}{\tl_use:N \g_yibin_declassify_year_tl}
\NewExpandableDocumentCommand{\yibinlogo}{}{\tl_use:N \g_yibin_logo_tl}

\NewDocumentCommand{\yibinsecretbox}{}{
  \bool_if:NTF \g_yibin_confidential_bool {$\boxtimes$} {$\square$}
}
\NewDocumentCommand{\yibinpublicbox}{}{
  \bool_if:NTF \g_yibin_confidential_bool {$\square$} {$\boxtimes$}
}
\ExplSyntaxOff

\AtBeginDocument{
  \hypersetup{
    pdftitle={\yibintitle},
    pdfauthor={\yibinauthor},
    pdfsubject={宜宾学院本科生毕业论文（设计） / \yibinenglishtitle},
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
  % The 2024 humanities Word template treats numbered level-one headings as
  % in-flow headings: left aligned, two-character indent, and no forced page.
  \titleclass{\chapter}{straight}
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
\renewcommand{\listfigurename}{插图目录}
\renewcommand{\listtablename}{表格目录}
\renewcommand{\cfttoctitlefont}{\hfill\yibinhei\bfseries\zihao{3}}
% Keep material after the trailing fill so TeX does not discard it at the
% paragraph boundary and accidentally right-align the title.
\renewcommand{\cftaftertoctitle}{\hfill\mbox{}}
\setlength{\cftbeforetoctitleskip}{0pt}
\setlength{\cftaftertoctitleskip}{24pt}
\renewcommand{\cftchapfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftchappagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsecfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsecpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsubsecfont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftsubsecpagefont}{\yibinsong\mdseries\zihao{-4}}
\renewcommand{\cftchapleader}{\cftdotfill{\cftdotsep}}
\setlength{\cftbeforechapskip}{0.25\baselineskip}
\ifyibin@science
  \setlength{\cftchapnumwidth}{2.5em}
  \setlength{\cftsecnumwidth}{3.5em}
  \setlength{\cftsubsecnumwidth}{4.5em}
\else
  \setlength{\cftchapnumwidth}{4em}
  \setlength{\cftsecnumwidth}{4em}
  \setlength{\cftsubsecnumwidth}{3.5em}
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

\newcommand{\yibin@coverfield}[3]{%
  \noindent
  \makebox[3.7cm][l]{\yibinhei\bfseries #1}%
  {\yibinsong\bfseries\underline{\parbox[c]{#2}{\centering\strut #3\strut}}}\par
  \vspace{0.35\baselineskip}
}

\newcommand{\yibin@coverdualfield}[6]{%
  \noindent
  \makebox[4.2cm][l]{\yibinhei\bfseries #1}%
  {\yibinsong\bfseries\underline{\parbox[c]{#2}{\centering\strut #3\strut}}}%
  \hspace{0.2cm}%
  \makebox[1.25cm][l]{\yibinhei\bfseries #4}%
  {\yibinsong\bfseries\underline{\parbox[c]{#5}{\centering\strut #6\strut}}}\par
  \vspace{0.35\baselineskip}
}

\NewDocumentCommand{\makeyibincover}{}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    \setlength{\parindent}{0pt}
    \begin{center}
      \vspace*{1.2cm}
      \IfFileExists{\yibinlogo}{%
        \includegraphics[width=10.5cm,height=3.0cm,keepaspectratio]{\yibinlogo}\par
      }{%
        {\yibinhei\bfseries\fontsize{34bp}{42bp}\selectfont 宜宾学院\par}
        \vspace{0.15cm}
        {\yibinenglishfont\bfseries\fontsize{18bp}{22bp}\selectfont YIBIN UNIVERSITY\par}
      }
      \vspace{0.8cm}
      {\yibinhei\bfseries\zihao{2} 本科生毕业论文（设计）\par}
      \vspace{1.2cm}
      \parbox{13.2cm}{\centering\yibinhei\bfseries\zihao{-2}\yibintitle\par}
      \vspace{1.3cm}
    \end{center}

    \begin{center}
      \begin{minipage}{13.3cm}
        \zihao{3}
        \yibin@coverfield{学院（部）}{9.2cm}{\yibincollege}
        \yibin@coverfield{专\hfill 业}{9.2cm}{\yibinmajor}
        \yibin@coverfield{学生姓名}{9.2cm}{\yibinauthor}
        \yibin@coverdualfield{学\hfill 号}{4.2cm}{\yibinstudentid}{年级}{3.2cm}{\yibingrade}
        \yibin@coverdualfield{指导教师（校内）}{4.2cm}{\yibinadvisor}{职称}{3.2cm}{\yibinadvisortitle}
        \yibin@coverdualfield{指导教师（校外）}{4.2cm}{\yibinexternaladvisor}{职称}{3.2cm}{\yibinexternaladvisortitle}
      \end{minipage}
    \end{center}
    \vspace{0.45cm}
    \begin{center}
      {\yibinsong\bfseries\zihao{3}\yibindate\par}
    \end{center}
  \endgroup
  \clearpage
}

\newcommand{\yibin@signatureline}[1]{%
  #1\hspace{0.5em}\underline{\makebox[4.2cm][c]{}}\ignorespaces
}

\newcommand{\makeyibinoriginality}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    \setlength{\parindent}{2em}
    \setstretch{2.0}
    \begin{center}
      {\yibinhei\bfseries\zihao{3} 原创性声明}
    \end{center}
    \vspace{1.6cm}

    本人呈交的学位论文（设计）《\yibintitle》，是在导师的指导下，独立进行研究取得的成果。
    除文中已经注明引用的内容外，本论文（设计）不包括其他个人或集体已经发表或撰写过的作品成果。
    对本文（设计）做出贡献的个人和集体，均已在文中以明确方式标明。
    本人完全意识到本声明的法律后果，因本声明而产生的法律后果由本人承担。

    \vspace{1.6cm}
    \noindent\yibin@signatureline{学位论文作者：}

    \vspace{0.8cm}
    \noindent 日期：\hspace{2em}年\hspace{2em}月\hspace{2em}日

    \vfill
    \noindent{\yibinhei\bfseries 附：《普通高等学校学生管理规定》（中华人民共和国教育部令第41号）}

    \noindent{\yibinhei\bfseries 第五十二条\quad 学生有下列情形之一，学校可以给予开除学籍处分：}

    \noindent{\yibinhei\bfseries （五）学位论文、公开发表的研究成果存在抄袭、篡改、伪造等学术不端行为，
    情节严重的，或者代写论文、买卖论文的；}
  \endgroup
  \clearpage
}

\newcommand{\makeyibincopyright}{%
  \clearpage
  \thispagestyle{empty}
  \begingroup
    \setlength{\parindent}{2em}
    \setstretch{2.0}
    \begin{center}
      {\yibinhei\bfseries\zihao{3} 学位论文（设计）版权使用授权书}
    \end{center}
    \vspace{1.4cm}

    本学位论文（设计）作者完全了解学校有关保留、使用学位论文（设计）的规定，
    同意学校保留并向国家有关部门或机构送交论文（设计）的复印件和电子版，
    允许论文（设计）被查阅和借阅。本人授权宜宾学院将本学位论文（设计）的全部或部分内容
    编入有关数据库进行检索，可以采用影印、缩印或扫描等复制手段保存和汇编。

    \vspace{1.0cm}
    \noindent 本学位论文（设计）属于（请在以下相应方框内打“√”）作品：

    \vspace{0.5cm}
    \noindent 保密\yibinsecretbox，在\underline{\makebox[2.0cm][c]{\yibindeclassifyyear}}年解密后适用本授权书。

    \vspace{0.3cm}
    \noindent 不保密\yibinpublicbox。

    \vspace{2.2cm}
    \noindent\yibin@signatureline{作者（签名）：}
    \hfill
    \yibin@signatureline{指导教师（签名）：}

    \vspace{0.9cm}
    \noindent 日期：\hspace{1.2em}年\hspace{1.2em}月\hspace{1.2em}日
    \hfill
    日期：\hspace{1.2em}年\hspace{1.2em}月\hspace{1.2em}日
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

% The current humanities sample keeps 绪论/结论 unnumbered, whereas the
% science hierarchy requires them to participate in chapter numbering.
\NewDocumentCommand{\yibinopeningchapter}{O{绪论}}{%
  \ifyibin@science\chapter{#1}\else\yibinunnumberedchapter{#1}\fi
}
\NewDocumentCommand{\yibinclosingchapter}{O{结论}}{%
  \ifyibin@science\chapter{#1}\else\yibinunnumberedchapter{#1}\fi
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

\endinput
