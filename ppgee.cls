% ------------------------------------------------------------------
% ------------------------------------------------------------------
% ppgee.cls v1.1
% Autoria: Gabriel Gutierrez P. Soares <https://github.com/gutierrezps>
% 
% Classe para elaboração de Dissertações e Teses do Programa de
% Pós-Graduação em Engenharia Elétrica do Instituto Federal da
% Paraíba (PPGEE-IFPB), Campus João Pessoa.
%
% Modelo elaborado de acordo com formatação do arquivo DOCX
% disponibilizado na página do PPGEE:
% https://www.ifpb.edu.br/ppgee/documentos/modelo-de-dissertacao/modelo_dissertacao_ppgee.docx
%
% PPGEE - IFPB: https://www.ifpb.edu.br/ppgee/
%
% Utilizando a classe abntex2 como base: https://www.abntex.net.br/
% e seu Modelo de Trabalho Acadêmico
%
% CHANGELOG
% * v1.1 - 2021/08/26
%   - Adoção de linguagem neutra nos comandos e textos de exemplo
%   - Adição de opção para especificar coorientação
%   - Atualização de links para página do PPGEE
% * v1.0 - 2021/04/21
%   - Versão inicial
% ------------------------------------------------------------------
% ------------------------------------------------------------------

\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{ppgee}[2021/08/26 PPGEE-IFPB Class v1.1 by gutierrezps]

% Booleanos e condições lógicas
\RequirePackage{ifthen}

% Usar opção "fonte=times" para usar a fonte Times New Roman
% Por padrão, a fonte Arial é usada
\newboolean{PPGEEfontetimes}
\setboolean{PPGEEfontetimes}{false}
\DeclareOption{fonte=times}{\setboolean{PPGEEfontetimes}{true}}

% Passa as demais opções para a classe abntex2
\DeclareOption*{%
  \PassOptionsToClass{\CurrentOption}{abntex2}%
}
\ProcessOptions\relax

% ---
% Carrega a classe abntex2 com opções pré-definidas
% ---
\LoadClass[
	% -- opções da classe memoir --
	12pt,				% tamanho da fonte
	openright,			% capítulos começam em pág ímpar (insere página vazia caso preciso)
	a4paper,			% tamanho do papel. 
	% -- opções da classe abntex2 --
	chapter=TITLE,		% títulos de capítulos convertidos em letras maiúsculas
	section=TITLE,		% títulos de seções convertidos em letras maiúsculas
	]{abntex2}
% ---

% Configuração das margens
\setlrmarginsandblock{2.54cm}{2.54cm}{*}
\setulmarginsandblock{2.54cm}{2.54cm}{*}
\checkandfixthelayout

% ---
% Seleção de fonte do documento
% ---
\ifthenelse{\boolean{PPGEEfontetimes}}{%
    % Fonte similar à Times New Roman
    \RequirePackage{newtxtext}%
    % Usar a fonte Times nos títulos dos capítulos, seções etc.
    \renewcommand{\ABNTEXchapterfont}{%
        \rmfamily\selectfont\bfseries%
    }%
}{% else
    % Fonte similar à Arial
    \RequirePackage{helvet}%
    \renewcommand{\familydefault}{\sfdefault}%
    % Títulos de capítulos, seções etc em negrito
    \renewcommand{\ABNTEXchapterfont}{\bfseries}%
    % Aumentar levemente fonte das equações
    \DeclareMathSizes{12}{14}{12}{10}
}
% ---

% Equações com a fonte Times New Roman
\RequirePackage{newtxmath}

% Recuo no primeiro parágrafo de cada seção.
\RequirePackage{indentfirst}

% Inclusão de gráficos
\RequirePackage{graphicx}

% Melhorias de justificação
\RequirePackage{microtype}

% Formatação de legendas de figuras, tabelas etc. (fonte menor, negrito)
\RequirePackage[font={footnotesize,bf}]{caption}


% Posiciona figuras e tabelas no topo da página quando adicionadas sozinhas
% em um página em branco. Ver https://github.com/abntex/abntex2/issues/170
\makeatletter
\setlength{\@fptop}{5pt} % Set distance from top of page to first float
\makeatother

% ---
% Formatação de parágrafos
% ---
% O tamanho do recuo da primeira linha é dado por:
\setlength{\parindent}{1.25cm}

% Controle do espaçamento entre um parágrafo e outro:
\setlength{\parskip}{0.2cm}
% ---

% ---
% Formatação de títulos de capítulos
% ---
\renewcommand{\ABNTEXchapterfontsize}{\large}
\setlength\afterchapskip{\onelineskip}

% Adicionar ponto após número do capítulo
\renewcommand{\printchapternum}{%
    \tocprintchapter
    \setboolean{abntex@innonumchapter}{false}
    \chapnumfont%
    \space\thechapter{}%
    \ifthenelse{\boolean{abntex@apendiceousecao}}{%
        \space
        \tocinnonumchapter
        \ABNTEXcaptiondelim%
    }{.\space} % else
}
% ---


% Formatação de títulos de seções
\renewcommand{\ABNTEXsectionfontsize}{\normalsize}
\renewcommand{\ABNTEXsectionfont}{\bfseries}
\setbeforesecskip{\onelineskip}
\setaftersecskip{\onelineskip}


% Formatação de títulos de subseção
\renewcommand{\ABNTEXsubsectionfontsize}{\normalsize}
\renewcommand{\ABNTEXsubsectionfont}{\bfseries}
\setbeforesubsecskip{\onelineskip}
\setaftersubsecskip{\onelineskip}


% Tradução dos labels para o comando autoref
\renewcommand{\tableautorefname}{Tabela}
\renewcommand{\figureautorefname}{Figura}

\renewcommand{\listadesiglasname}{Lista de siglas}

% ---
% Lista de símbolos
% ---
\renewenvironment{simbolos}
    {%
        \pretextualchapter{\listadesimbolosname}
        \list{}% empty label
            {\setlength{\topsep}{0pt}
            \setlength{\partopsep}{0pt}
            \setlength{\itemsep}{0pt}
            \setlength{\parsep}{0pt}
            \setlength{\leftmargin}{2cm}
            \setlength{\rightmargin}{0em}
            \setlength{\listparindent}{1em}
            \setlength{\itemindent}{0em}
            \setlength{\labelwidth}{4em}
            \setlength{\labelsep}{1em}}%
    }%
    {\endlist\cleardoublepage}
% ---

% Ambiente para configurar espaçamento entre linhas
\newenvironment{espacamento}[1]%
    {\let\oldespacamento\baselinestretch
        \linespread{#1}\selectfont%
    }{\linespread{\oldespacamento}\selectfont\par}


% Configurações da folha de aprovação
\setlength{\ABNTEXsignskip}{0.5cm}      % espaço entre assinaturas
\setlength{\ABNTEXsignwidth}{15cm}      % largura das assinaturas
\setlength{\ABNTEXsignthickness}{0pt}   % linha de assinatura (removida)


% numeração contínua de equações
\counterwithout{equation}{chapter}

% títulos das referências em negrito (o padrão é em itálico)
\AtBeginDocument{\citeoption{abnt-emphasize=bf}}

% Retira espaço extra obsoleto entre as frases.
\AtBeginDocument{\frenchspacing}

% Nomes constantes da classe ppgee em portugues (definicoes padroes)
\newcommand{\orientacaoname}{Orientação}
\newcommand{\coorientacaoname}{Coorientação}

\addto\captionsenglish{% ingles
    %% adjusts names from abnTeX2
    \renewcommand{\orientacaoname}{Supervisor}
    \renewcommand{\coorientacaoname}{Co-supervisor}
}

% Comandos de dados - autoria (use \and para multiplos autores)
\newcommand{\autoria}[1]{\author{#1}}
\newcommand{\imprimirautoria}{\theauthor}


% Comandos de dados - email
\providecommand{\imprimiremailAutoria}{}
\newcommand{\emailAutoria}[1]{\renewcommand{\imprimiremailAutoria}{#1}}

% Comandos de dados - areaConcentracao
\providecommand{\imprimirareaConcentracao}{}
\newcommand{\areaConcentracao}[1]{\renewcommand{\imprimirareaConcentracao}{#1}}

% Comandos de dados - orientacao
\providecommand{\imprimirorientacaoRotulo}{}
\providecommand{\imprimirorientacao}{}
\newcommand{\orientacao}[2][\orientacaoname]%
  {\renewcommand{\imprimirorientacaoRotulo}{#1}%
   \renewcommand{\imprimirorientacao}{#2}}

% Comandos de dados - coorientacao
\providecommand{\imprimircoorientacaoRotulo}{}
\providecommand{\imprimircoorientacao}{}
\newcommand{\coorientacao}[2][\coorientacaoname]%
  {\renewcommand{\imprimircoorientacaoRotulo}{#1}%
   \renewcommand{\imprimircoorientacao}{#2}}


% ---
% Conteúdo padrão da Capa
% ---
\renewcommand{\imprimircapa}{\begin{capa}%
\begin{center}
    %\includegraphics[width=\textwidt]{img/ifpb_guarabira.jpg}
    \includegraphics[width=0.5\textwidth]{img/ifpb_guarabira.jpg}    
    \vspace{1.5cm}
    
    \textbf{\large\imprimirautoria}
    
    \vspace{6cm}
    
    
    \textbf{\Large\imprimirtitulo}
    
    
    \vfill
    
    \textbf{\imprimirlocal}\\
    \textbf{\imprimirdata}
\end{center}
\end{capa}}
% ---

% ---
% Conteúdo padrão da Folha de Rosto
% ---
\makeatletter
\renewcommand{\folhaderostocontent}{%
  \begin{center}
    %\includegraphics[width=\textwidt]{img/ifpb_guarabira.jpg}
    \includegraphics[width=0.5\textwidth]{img/ifpb_guarabira.jpg}
    
    \vspace{1.5cm}
    
    \textbf{\large\imprimirautoria}
    
    \vspace{1.5cm}
    
    \textbf{\Large\imprimirtitulo}
    
    \vfill
    
    \hfill
    \begin{minipage}{.5\textwidth}
        \SingleSpacing
        \imprimirpreambulo
    \end{minipage}%
    
    \vfill
    
    Área de Concentração: \imprimirareaConcentracao
    
    \vspace{1.5cm}
    
    \imprimirorientacao\\
    \imprimirorientacaoRotulo
    
    \vspace{1cm}
    
    \imprimircoorientacao\\
    \imprimircoorientacaoRotulo
    
    \vspace{1cm}
    
    \imprimirlocal, \imprimirdata\\
    \textcopyright~\imprimirautoria~ --- \imprimiremailAutoria
  \end{center}
}
\makeatother