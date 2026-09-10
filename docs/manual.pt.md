# Manual do LynxAutomator

[Español](manual.md) · [English](manual.en.md) · [Instalação e pacotes](../README.md#linux-macos-and-windows-builds)

Este manual descreve a aplicação da branch `feature/qt-migration`, ainda com interface Tk.
Os executáveis antigos publicados em Releases podem ter um comportamento diferente.
O [guia DOCX original](../WIP%20LynxAutomator%20GUIDE%20.docx) e as suas capturas são
conservados como referência histórica; as instruções atuais são mantidas em Markdown.

## 1. Preparação

O LynxAutomator prepara ficheiros Excel, descarrega imagens autorizadas do Wildlife
Insights e disponibiliza ferramentas de armadilhagem fotográfica. Não envia
automaticamente o Excel nem as imagens para o Wildbook.

Pode escolher português, espanhol ou inglês na interface. Alguns nomes de separadores,
fases de processamento e mensagens de erro permanecem em inglês. Guarde os resultados
antes de mudar o idioma, pois essa ação reconstrói os formulários.

Mantém-se uma única aplicação para Windows, Linux e macOS, com Bulk Import,
descargas, acompanhamento do lince, vídeo, correção de datas e renomeação.
A edição alpha mini foi retirada desta branch e permanece no histórico.
A migração para Qt parte desta aplicação e dos seus serviços partilhados.
Consulte o [plano da interface](interface-plan.md).

Siga o README para obter e extrair o pacote do seu sistema. Os novos pacotes têm de
passar pelas verificações de GitHub Actions. Ao corrigir datas ou renomear, comece
por cópias dos ficheiros originais.

### Tarefas longas, progresso e cancelamento

Vídeos, leitura e criação de Excel, catálogos, acompanhamento do lince, leitura e
correção de datas, renomeação e descargas são processados em segundo plano. A barra
inferior mostra a tarefa, a fase ou o ficheiro atual e o botão **Cancelar**. Quando
a duração não é conhecida, a barra indica atividade em vez de uma percentagem.

Durante uma tarefa, os campos e botões de operação ficam temporariamente desativados
para impedir alterações aos parâmetros ou duas operações sobre os mesmos ficheiros.
A janela continua a responder. Aguarde a conclusão ou cancele antes de iniciar outra
tarefa ou mudar de idioma.

O cancelamento é cooperativo: verifica-se entre fotogramas, ficheiros ou etapas de
processamento. Uma leitura/escrita de Excel, cópia, chamada ao descodificador ou
transferência já iniciada pode precisar de terminar primeiro. Uma transferência
gsutil tem um limite de cinco minutos. Não é uma interrupção instantânea.

Os ficheiros já concluídos são preservados. Cancelar não desfaz renomeações nem
correções já aplicadas aos originais. Os fotogramas e as cópias em preparação usam
ficheiros temporários; um Excel só substitui o destino depois de estar totalmente
escrito e de se verificar o cancelamento. Ao fechar a janela durante uma tarefa,
a aplicação pede o cancelamento e aguarda o fim do trabalhador antes de fechar.

## 2. Excel inicial do Wildbook

Use o modelo `.xlsx` do seu projeto, com os nomes exatos das colunas esperadas pela
sua instância do Wildbook. Este exemplo inclui apenas alguns campos comuns:

| Encounter.locationID | Encounter.country | Encounter.genus | Encounter.specificEpithet | Encounter.submitterID |
| --- | --- | --- | --- | --- |
| Andújar-Cardeña | Spain | Lynx | pardinus | o_seu_utilizador |

Inclua uma primeira linha com os valores comuns. Os módulos podem propagar esses
valores. Deixe vazios os campos de fotografias, datas e indivíduos que devam ser
calculados; confirme as coordenadas e os restantes requisitos do projeto.

A saída é uma tabela de dados: não é garantida a conservação da formatação,
fórmulas ou folhas adicionais do livro original. Guarde o resultado com um nome
diferente do modelo. A própria gravação do Excel também decorre em segundo plano.

## 3. Wildbook → BIWbE a partir de pasta

1. Selecione a pasta que contém diretamente as imagens PNG/JPG/JPEG.
2. Selecione o Excel inicial.
3. Se pretender várias imagens por encontro, ative a opção de agrupamento e indique
   um limiar inteiro não negativo em **segundos**.
4. Processe e, no fim, guarde o Excel através do botão de descarga.

Este módulo não percorre subpastas. Só inclui imagens com uma data EXIF
`DateTimeOriginal` legível; não usa a data do ficheiro como substituição. Verifique
o número de imagens incluídas.

O agrupamento compara imagens consecutivas ordenadas por data. Com um limiar de
60 segundos, imagens às 10:00:00, 10:00:40 e 10:01:20 podem ficar no mesmo encontro.
Neste módulo não há separação por projeto ou câmara: use uma pasta coerente por
câmara ou instalação.

## 4. Wildbook → Catálogo BIWbE

1. Selecione a pasta do catálogo; neste módulo as subpastas são percorridas.
2. Selecione o Excel inicial.
3. Escolha se pretende capitalizar o identificador e juntar linhas do mesmo indivíduo.
4. Processe e guarde o resultado.

A primeira palavra do nome do ficheiro, separada por espaços, torna-se
`MarkedIndividual.individualID`. Por exemplo, `Nube lateral.jpg` identifica `Nube`,
enquanto `Nube_lateral.jpg` identifica `Nube_lateral`. A capitalização pode alterar
as restantes letras do identificador; desative-a se isso afetar os seus dados.
As referências às imagens são caminhos relativos à pasta selecionada.

Se o identificador existir apenas no nome da pasta, pode preparar cópias com o
renomeador. Conserve o espaço que separa o indivíduo do resto do nome; substituí-lo
por um sublinhado altera a interpretação feita pelo catálogo.

## 5. Bulk Import → Wildlife Insights: fotos locais ou descargas

As descargas estão integradas em **Bulk Import**, a primeira aba. Carregue o ZIP WI, selecione espécies e reveja a seleção. Escolha **Fotos locais** para pesquisar imagens numa pasta e subpastas, ou **Descarregar fotografias** para obter referências `gs://` com gsutil. Clique em **Usar fotos locais** ou **Descarregar fotos selecionadas**, conforme o modo escolhido e depois em **Configurar Excel**, quando houver fotos verificadas.

Instale Google Cloud CLI com gsutil. Use **Autorização (opcional) → Iniciar sessão com Google** e conclua o acesso no navegador com uma conta autorizada. Google Cloud CLI gere e conserva as credenciais; LynxAutomator não pede a sua palavra-passe Google. Iniciar sessão não concede novas permissões no bucket. Para gsutil independente ou configurações personalizadas, siga as instruções fornecidas na exportação WI.

O conteúdo é verificado e as extensões preservadas. Cada descarga cria um novo lote com `manifest.csv`, sem sobrescrever fotos existentes. Pode repetir as falhas. Cancelar espera pela transferência atual (máximo cinco minutos) e conserva os ficheiros concluídos.

## 6. Wildlife Insights–Wildbook → CSVs WI para BIWbE

Selecione três ficheiros: o Excel inicial (primeira folha), o CSV de imagens e o
CSV de instalações. Ambos os CSV precisam de `project_id` e `deployment_id`.
Depois da união, devem estar disponíveis `latitude`, `longitude`, `placename`,
`location`, `timestamp`, `project_id`, `deployment_id` e `subproject_name`.
Se a coluna `number_of_objects` não existir, assume-se o valor 1.
Colunas de dados com o mesmo nome nos dois CSV podem provocar ambiguidades;
confirme os cabeçalhos se surgir um erro.

Escolha as opções, processe, resolva os erros apresentados e guarde o Excel:

- **Várias imagens por encontro**: agrupa imagens consecutivas pelo limiar inteiro
  não negativo em segundos, separadamente por projeto e instalação.
- **Separar imagens com mais de um objeto**: coloca cada imagem com
  `number_of_objects > 1` numa linha própria; não cria uma linha por animal.

**Estas opções não significam “multiespécies”.** O código não identifica nem compara
espécies. Com ambas ativas, as imagens com mais de um objeto ficam separadas e as
restantes podem ser agrupadas por tempo.

Identificadores vazios, instalações duplicadas, imagens sem instalação correspondente
e datas inválidas são rejeitados. Corrija as entradas e volte a processar. Uma
falha no processamento não disponibiliza o Excel de uma execução anterior.

As referências usam o nome base com extensão `.JPG`. Confirme a correspondência com
as imagens a carregar, especialmente se o descarregador tiver substituído caracteres
ou se existirem nomes iguais entre instalações. `Occurrence.occurrenceID` continua
a ser `projeto-instalação`, não um identificador único por sequência; confirme a
adequação ao esquema de importação do seu projeto.

## 7. Módulo do lince ibérico

Selecione **a pasta que contém as propriedades (Finca)**. Ajuste as opções de revisão
e pasta intermédia Linces de acordo com a estrutura real:

| Pasta Revisión | Pasta Linces | Caminho dentro da pasta selecionada |
| --- | --- | --- |
| Não | Não | `Finca/Estación/Individuo/imagem.jpg` |
| Sim | Não | `Finca/Estación/Revisión/Individuo/imagem.jpg` |
| Não | Sim | `Finca/Estación/Linces/Individuo/imagem.jpg` |
| Sim | Sim | `Finca/Estación/Revisión/Linces/Individuo/imagem.jpg` |

1. Escolha a pasta raiz em **Source**.
2. Em **Settings**, ajuste as duas opções e indique os minutos de agrupamento:
   `0` desativa-o; um inteiro positivo agrupa registos próximos no tempo.
3. Em **Optional Files**, pode selecionar um Excel de estações com coluna `Estacion`
   e outro de indivíduos com coluna `Lince`. Evite chaves duplicadas para não
   multiplicar registos nas uniões.
4. Gere, reveja e guarde o Excel.

Uma pasta de indivíduo chamada `Nube y Brisa` cria registos para ambos. A conjunção
reconhecida pelo programa é ` y ` (ou ` Y `), mesmo na interface portuguesa.
O agrupamento reúne ficheiros e indivíduos por propriedade, estação e revisão.

Os caminhos são separados por `;`. `Número de Fotos` conta essas entradas, que podem
incluir vídeos ou repetições quando há vários indivíduos; não é uma contagem validada
de fotografias únicas. As datas provêm de EXIF quando legível; os ficheiros sem essa
informação, incluindo habitualmente vídeos, podem ficar sem data. Reveja-os antes
de usar agrupamento temporal.

Os campos dependem das posições das pastas. Níveis adicionais ou opções incorretas
podem atribuir nomes aos campos errados; confirme-os no resultado.

## 8. Alteração de datas

Aplica o mesmo desvio do relógio aos ficheiros diretamente na pasta, sem subpastas.

1. Selecione a pasta e aguarde a leitura das datas em segundo plano.
2. Escolha a data mais antiga, mais recente ou uma referência personalizada que
   represente a hora incorreta da câmara.
3. Introduza a hora real correspondente no formato `AAAA-MM-DD HH:MM:SS`.
4. Escolha copiar para outra pasta ou reescrever as datas dos originais.

Se a câmara indicava `2024-05-01 10:00:00` quando eram 12:00, o desvio é +2 horas
para todos os ficheiros; não ficam todos com a mesma data.

| Dados | Alteração |
| --- | --- |
| JPEG com EXIF DateTimeOriginal | Datas de captura, digitalização e DateTime recebem o desvio |
| Outros formatos/JPEG sem data de captura | Não é criada nem corrigida uma data de captura incorporada |
| Datas do ficheiro no Windows | Criação, modificação e acesso |
| Datas do ficheiro no Linux/macOS | Modificação e acesso; não criação |

A referência mais antiga/recente usa EXIF legível ou, na sua ausência, criação no
Windows e modificação no Linux/macOS. Não existe conversão automática de fuso
horário. Reveja a referência se misturar formatos ou câmaras.

As cópias usam a referência do original e nomes alternativos quando já existe um
destino. Um lote pode terminar parcialmente; cancelar preserva o que já foi
concluído. Após uma alteração dos originais, as datas são relidas, incluindo após
cancelamento ou erro. Verifique o novo desvio antes de repetir uma correção.

## 9. Extração de fotogramas de vídeo

1. Selecione a pasta que contém diretamente vídeos MP4/AVI/MOV/MKV/FLV.
2. Introduza um intervalo positivo e finito em **segundos**, por exemplo `1`.
3. Inicie a extração e selecione a pasta de saída.
4. Acompanhe o progresso na barra inferior; pode cancelar entre fotogramas.

As subpastas não são percorridas. Os fotogramas são JPEG, extraídos em intervalos
aproximados a um número inteiro de fotogramas, com um mínimo de um fotograma.
Os nomes incluem o vídeo e um número; são acrescentados sufixos para preservar
ficheiros existentes. Uma falha num vídeo é apresentada no resumo, sem impedir o
processamento dos restantes. Cancelar liberta o vídeo e limpa a saída temporária.

Antes da extração, **ffprobe** (do FFmpeg, disponível no PATH) lê as datas internas.
Reveja os valores e a origem de cada vídeo; confirme ou corrija usando a hora
impressa pela câmara. Sem metadados, com conflitos ou sem ffprobe, introduza a data
manualmente. Nunca se usa automaticamente a data do ficheiro.

Formato: `2024-07-15 14:30:00`, opcionalmente com desvio UTC como `+02:00`.
Sem desvio, o EXIF mantém a hora da câmara e as datas do ficheiro não são alteradas.
Com desvio, são gravadas as etiquetas EXIF de zona e ajustados acesso/modificação;
o Windows também ajusta criação com precisão de milissegundos, Linux/macOS não.
Cada fotograma soma a sua posição (número/FPS). São gravadas captura, digitalização,
modificação e frações de segundo EXIF. Com FPS variável, o tempo é aproximado.
A hora impressa não é lida automaticamente por OCR.

## 10. Renomeação de imagens

1. Selecione a origem; são percorridas as suas subpastas para JPG/JPEG/PNG/GIF.
2. Escolha os componentes: primeira palavra da pasta, nome original, data EXIF
   e/ou texto personalizado. Pode substituir espaços por sublinhados.
3. Ative a cópia e escolha um destino diferente para conservar os originais.
   Caso contrário, os ficheiros são renomeados na sua localização atual.
4. Inicie e acompanhe o progresso; cancele entre ficheiros se necessário.

As opções de pasta e nome original aplicam capitalização. O texto personalizado
não admite `/` nem `\`. Se não existir data EXIF, esse componente é omitido.
Ao copiar, as imagens são reunidas numa única pasta, sem recriar a árvore original.
Colisões recebem sufixos. Uma pasta de destino dentro da origem é excluída do
percurso para não voltar a processar as cópias.

Cancelar não desfaz as alterações já terminadas. Reveja o resumo de ficheiros
concluídos, ignorados e falhados.

## 11. Problemas frequentes e limites

| Situação | O que verificar |
| --- | --- |
| Aviso de editor desconhecido ou bloqueio ao abrir | Origem do pacote e política do computador; Não desative a proteção para experimentar ao acaso. |
| Falta Tkinter | Instalação de Python com Tk, conforme o README. |
| Falta gsutil ou acesso negado | Google Cloud CLI, PATH, conta autorizada e guia incluído na exportação WI. |
| CSV rejeitado | Colunas, identificadores, instalações duplicadas/não correspondentes e datas. |
| Faltam imagens no BIWbE a partir de pasta | EXIF de captura e seleção de uma pasta com as imagens diretamente dentro dela. |
| Não é possível guardar Excel | Feche o livro no Excel e escolha um destino com permissão de escrita. |
| A indicação de cancelamento permanece visível | Aguarde a conclusão da operação em curso; cancelar não interrompe à força uma escrita ou chamada externa. |
| Os botões estão desativados | Há uma tarefa em curso; use Cancelar ou aguarde o resultado. |

Guarde antes de fechar ou mudar de idioma. Não existe um histórico universal de
anulação. Criar o Excel não valida todas as regras do Wildbook nem a integridade
do conjunto de dados. Consulte a [revisão de lógica](logic-review.md).

## Fotografias Camtrap DP

Abra **Bulk Import → Camtrap DP**, carregue o ZIP completo (1.x, CSV/CSV.gz) e
marque as espécies usando a pesquisa e as caixas de seleção. Não é necessário
registar um Wildbook. Reveja o número de imagens únicas antes de escolher o destino.
As imagens associadas por evento são incluídas por predefinição, usando a instalação
e o intervalo temporal, mesmo vazias, para deteção no Wildbook. Isso não significa
que todos os animais sejam o mesmo indivíduo. O modo local não acede à Internet.

A pasta `camtrap-…` contém as imagens verificadas e `manifest.csv` com IDs, espécies,
nomes e estados. Vídeos são omitidos. Fotos privadas podem exigir credenciais
configuradas em **Autorização (opcional)**. Repetir falhadas cria um novo lote.
Cancelar preserva os ficheiros completos e limpa os parciais; uma leitura HTTP em
curso pode esperar até ao seu limite de 20 segundos.

O exemplo sintético local opcional (excluído do Git) tem 366 observações de lince, 247 imagens diretamente
associadas, 300 incluindo eventos e 10 JPEG locais. As fotos não mostram linces.
A validação é estrutural básica, não completa. Depois de preparar as fotografias, **Configurar Excel** abre o editor comum. A carga direta de projetos está disponível nas origens API (alpha).

## Bulk Import comum: Wildlife Insights e Camtrap DP

Em **Bulk Import → Wildlife Insights**, carregue apenas o ZIP exportado, com `images.csv` ou `images_<projeto>.csv` e um único `deployments.csv` na mesma pasta do ZIP. As espécies são lidas automaticamente. Selecione as espécies e clique em **Rever seleção**, depois em **Usar fotos locais** ou **Descarregar fotos selecionadas**, conforme o modo escolhido. Pode descarregar referências `gs://` com gsutil instalado ou ativar **Fotos locais**, que também pesquisa subpastas. Sem gsutil, a opção local está inicialmente selecionada. Os nomes devem corresponder a `location`; nomes ambíguos e imagens inválidas são indicados como falhas. As descargas usam uma nova pasta com `manifest.csv`, preservando os ficheiros existentes. Pode repetir as falhas. **Configurar Excel** fica disponível quando existem fotos verificadas; a pré-visualização informa as pendentes e exporta apenas as disponíveis. Não é necessário um modelo Excel.

O modelo Excel antigo continua nas opções avançadas, por compatibilidade. Excel é o formato de saída. No Camtrap DP, selecione espécies e obtenha as fotos antes de abrir Bulk Import; os lotes e novas tentativas são reunidos durante a sessão.

Edite, adicione, desative, remova e ordene campos com ↑. `fixed` usa um valor constante; os outros campos usam dados do registo. Use nomes aceites pelo Wildbook. Pode agrupar por evento/espécie/indivíduo em Camtrap ou por intervalo/projeto/instalação/espécie em WI. Fotos com vários animais ficam separadas. Clique em **Pré-visualizar** e **Guardar Excel**. Não há envio automático.

### Perfis opcionais por localidade

Escreva um nome, como «Doñana», e clique em **Guardar perfil**. Use **Carregar perfil** para reutilizar localização, país, remetente e campos adicionais. Guardar o mesmo nome atualiza esse perfil. Pode trabalhar sem guardar; nenhum perfil é carregado automaticamente e a pré-visualização não guarda alterações.

Os perfis são partilhados entre as fontes e ficam em `.local-settings/bulk-import.json`, ignorado pelo Git. O perfil anterior aparece como `Default`. Na aplicação empacotada ficam em `LynxAutomator` dentro de `LOCALAPPDATA` ou `XDG_CONFIG_HOME`/`~/.config`.

A validação é local, não verifica o servidor. As datas mantêm a hora da origem. Agrupar eventos não demonstra identidade individual.

### Catálogo e comentários com metadados

O ZIP de Wildlife Insights aceita `images.csv` e `images_<projeto>.csv`, combinando fragmentos da mesma pasta com um único `deployments.csv`. Conjuntos em pastas diferentes são rejeitados para evitar misturas.

Os nomes das colunas têm uma lista editável baseada na [documentação do Wildbook](https://wildbook.docs.wildme.org/data/bulk-import-beta.html), incluindo campos de encontros, avistamentos, projetos e outros. Pode escrever nomes personalizados ou alterar índices; os campos de fotografias são automáticos, mas subcampos como `.keywords` são configuráveis. A disponibilidade depende do servidor.

Use **Adicionar campo → Comentários com metadados** para pesquisar e marcar vários campos dos dados. Escolha `Sighting.comments`, `Encounter.sightingRemarks` ou `Encounter.researcherComments` e adicione-os aos comentários. O texto anterior é preservado. Exemplo de origem `template`:

```text
Câmara: {deployment.cameraID}; Modelo: {deployment.cameraModel}; Instalação: {deployment.setupBy}
```

Os campos disponíveis usam os prefixos `deployment.`, `media.` e `observation.`. Valores distintos são preservados ao agrupar, separados por ` | `. Campos inexistentes produzem um erro explícito. As configurações ficam nos perfis locais. As notas não substituem a conservação dos CSV originais.

Perfis novos usam `Encounter.sightingID` para ligar o avistamento; `Encounter.sightingRemarks` preserva comentários em encontros clonados. A validação identifica campos obrigatórios ausentes: género, espécie, ano, foto e localização (texto, locationID ou ambas as coordenadas). Comentários maiores que o limite Excel de 32767 caracteres são rejeitados, nunca truncados silenciosamente.

### Configuração comum e metadados adicionais

WI e Camtrap DP partilham editor, validação, perfis, agrupamento e Excel. Defina o intervalo máximo em segundos no editor ao ativar o agrupamento. Eventos explícitos de Camtrap são preservados; séries sem evento usam o intervalo. Alterações exigem nova pré-visualização.

Os CSV adicionais do ZIP, como `projects.csv`, são lidos automaticamente. São associados por identificadores de projeto, câmara, instalação ou imagem. Uma tabela global de uma linha sem identificadores pode fornecer valores comuns. A pré-visualização avisa quando não há correspondência. PDFs não são convertidos em campos. Camtrap também disponibiliza `package.*` e recursos CSV adicionais declarados. Use os campos em comentários, por exemplo `{projects.project_name}` ou `{cameras.camera_model}`.

### Wildbook e seleção múltipla

O seletor normal mostra apenas nomes de Wildbooks. **Opções avançadas → Consultar ramos GitHub** obtém a lista atual de ramos para selecionar pelo nome, sem mostrar URLs. O JSON local fica nas opções avançadas; listas e catálogos são guardados em cache. Um ramo de desenvolvimento pode não conter catálogo válido.

Selecione várias instalações ou encontros com **Ctrl/Shift** e aplique a mesma localização. Encontro tem prioridade sobre instalação e valor comum; `*` aplica a todas as linhas e elimina exceções. Guarde no perfil. Ao mudar o agrupamento, reveja atribuições por encontro. As coordenadas são preservadas; GitHub pode diferir do servidor. As janelas ficam associadas à janela de origem para aparecer à frente.

### Seleção estável e fluxo por localizações

As listas de seleção permanecem abertas e incluem pesquisa; escolha uma linha ou feche com Escape. No Linux, Zenity é utilizado para abrir, guardar e escolher pastas. Instale `zenity` se necessário.

Em WI: **carregar ZIP → selecionar espécies → rever seleção → obter fotografias → configurar Excel**. WI e DP partilham os controlos de seleção, aquisição e acesso ao editor, com adaptadores específicos para cada formato.

O seletor locationID começa por localidades de origem; também permite coordenadas e campos de país/localização disponíveis. Selecione várias com Ctrl/Shift. Ao aplicar, mostra uma marca e o número de linhas afetadas. Ramos avançados ficam no fundo.

Pasta e catálogo usam o editor comum. Indique espécie, subpastas e método de identidade explícito (nenhum, primeira palavra do ficheiro ou nome da pasta). Datas vêm do EXIF; pode indicar apenas um ano quando falta a data, sem inventar mês/dia nem agrupar essas fotos por tempo. Sem EXIF nem ano, as fotos omitidas são indicadas. Complete a localização no editor. O modelo anterior continua nas opções de compatibilidade.


Bulk Import reúne Wildlife Insights, Camtrap DP, Criar a partir de pasta, Catálogo, Agouti API e Trapper API na primeira aba; **Configurar Excel** abre uma janela comum. Selecione a origem, prepare os dados e configure os campos, a pré-visualização e o Excel. O seletor de origem permite rever a fonte atual ou escolher outra; selecionar uma origem descarta o editor atual. Catálogo aceita fotos sem data, com campos temporais vazios e sem agrupamento temporal. Lince Ibérico está em Funcionalidades.


### Acesso Camtrap DP a fotografias privadas

Em **Bulk Import → Camtrap DP**, **Fotos locais** pede primeiro a pasta de originais e depois o destino das cópias preparadas. Referências remotas podem ser associadas por nome; correspondências ambíguas são rejeitadas. **Descarregar fotografias** obtém as URL do pacote e copia imagens locais incluídas. Os vídeos continuam excluídos.

**Autorização (opcional)** aceita **Agouti API key**, **Agouti Bearer** ou **Trapper token**. Obtenha a chave ou token Agouti da sua conta. Para Trapper, abra o servidor, inicie sessão e gere um token API no perfil. Indique o servidor HTTPS que exige a credencial. As credenciais ficam apenas na memória, nunca nos perfis, Excel ou manifestos; **Remover acesso da memória** elimina-as da aplicação. As credenciais não são reenviadas para outro servidor nem para HTTP.

Fotos privadas podem ser solicitadas no modo de descarga; o servidor verifica as permissões da conta ou ligação. Erros 401/403 indicam que deve configurar acesso ou rever permissões, podendo depois repetir as falhas. Configurar um token não valida permissões até à descarga. Para carregar diretamente os dados de um projeto, use as origens Agouti API ou Trapper API.


Referencias / References: [Agouti](https://docs.agouti.eu/api/endpoints.html), [Trapper](https://trapper-project.readthedocs.io/en/latest/tutorial.html#authentication), [Google Cloud CLI](https://cloud.google.com/sdk/docs/authorizing).

O intervalo em segundos fica junto a **Agrupar fotografias**. **Escolher locationID** aparece apenas na linha `Encounter.locationID`. Os nomes oficiais das colunas são preservados; os nomes das variáveis de origem não são traduzidos e os tipos de dados usam o idioma selecionado, sem alterar os perfis guardados.


### Origens Agouti API e Trapper API

**Wildlife Insights** é a origem inicial. **Criar a partir de pasta** substitui Pasta. WI, DP e fontes API partilham seleção de espécies, revisão, preparação de fotos e configuração do Excel. **Fotos locais** é a opção inicial; também pode escolher descarregar. Configure **Autorização (opcional)** apenas se o servidor exigir credenciais.

Use **Consultar valores do projeto** e depois **Selecionar dados…** para escolher valores antes de **Carregar seleção de dados**. Agouti lê descritor/instalações antes de pedir media/observações; Trapper lê o ZIP completo de metadados e reutiliza-o ao alterar filtros. Nenhum destes passos pede fotografias. A autorização permanece em memória e restrita ao servidor configurado.

### Exportação e filtros

**Pré-visualizar** valida e abre uma janela com até 100 linhas; o Excel inclui todas. Alterar campos ou agrupamento exige nova pré-visualização. `MarkedIndividual.individualID` começa desmarcado; os perfis existentes conservam a seleção. **Obter valor de** carrega os metadados ao abrir, incluindo relações por IDs entre instalação, câmara e modelo. Partilhar apenas o projeto não identifica uma câmara. **Adicionar campo** oferece coluna ou comentários com seleção de metadados visíveis.

Em **Escolher locationID**, pode agrupar por qualquer metadado disponível, selecionar várias linhas com Ctrl/Shift ou todas e aplicar várias atribuições. **Fechar** regressa ao editor. O perfil guarda a hierarquia. O nome sugerido ao guardar inclui o ID do ancestral comum mais específico e data/hora local (`AAAA-MM-DD_HH-MM-SS`), sem alterar os IDs nas linhas. Sem hierarquia conhecida usa `varias-ubicaciones`; IDs ausentes usam `sin-ubicacion`/`ubicaciones-incompletas`. O nome é editável.

**Selecionar dados…** oferece multiseleção com pesquisa, valores existentes, contagens e número de instalações correspondentes. Conforme as colunas presentes: `deploymentStart.year`, `locationName`, `locationID`, `deploymentID`, `cameraID`, `cameraModel`, `habitat`, `setupBy`, `deploymentTags`, `captureMethod` e `baitUse`. Colunas ausentes não aparecem. Valores da mesma variável são alternativas; variáveis diferentes intersectam-se. Sem marcas aceita qualquer valor. Resultados vazios impedem continuar. Ano significa início da instalação, não recorte temporal de fotografias. As espécies selecionam-se depois.

Agouti pede media/observações por IDs escolhidos, reutilizando o índice de instalações. No Trapper a seleção é local depois de descarregar o ZIP de metadados e não reduz essa transferência. Ambas as origens permanecem **alpha**, sem testes com contas reais. No WI, inclua filtros de Catalogued/Identify ao pedir o ZIP e carregue-o aqui. As fotos só se obtêm com **Usar fotos locais** ou **Descarregar fotos selecionadas**.

### Níveis de pastas e estações

**Criar a partir de pasta** e **Catálogo** abrem a mesma janela **Configurar Excel**, com `Encounter.locationID`, `Encounter.submitterID`, perfis e comentários. Fechá-la preserva a página de origem.

Ative **Interpretar os níveis das subpastas**. Nível 1 é a primeira pasta sob a raiz; `—` não atribui esse dado. Para `Fotos/Localidade/Estação/Indivíduo/foto.jpg`, selecione Fotos como raiz, localidade=1, estação=2, indivíduo=3. Pode atribuir um nível ao nome científico quando as pastas contêm nomes binomiais; caso contrário escreva um nome científico comum. O nível de indivíduo tem prioridade sobre o nome do ficheiro; ative `MarkedIndividual.individualID` no editor se o quiser exportar.

**Rever subpastas e estações** enumera caminhos de estações. Edite a localidade e ambas as coordenadas WGS84 decimais, ou deixe ambas vazias; zero é válido. **Guardar estações** aplica os valores às fotos correspondentes. Estações homónimas sob localidades diferentes continuam distintas. Não é necessário criar instalações. Níveis ausentes são indicados; alterar níveis/raiz exige nova revisão. A tabela mantém-se durante a sessão. Pasta exige EXIF ou ano de recurso; Catálogo aceita fotos sem data.

**Obter valor de** mostra nomes de variáveis sem tradução. WI oferece `images.*`, `deployments.*`, `projects.*`, `cameras.*` e os CSV relacionados, incluindo tabelas partidas por projeto. DP oferece os seus recursos; pastas usam `folder.*`, `station.*`, `file.*`. As relações seguem IDs explícitos sem misturar projetos. Os aliases antigos de perfis WI continuam legíveis, mas não aparecem como novas opções.
