# Manual do LynxAutomator

[Español](manual.md) · [English](manual.en.md) · [Instalação e pacotes](../README.md#linux-macos-and-windows-builds)

Este manual descreve a aplicação completa da branch `feature/cross-platform-builds`.
Os executáveis antigos publicados em Releases podem ter um comportamento diferente.
O [guia DOCX original](../WIP%20LynxAutomator%20GUIDE%20.docx) e as suas capturas são
conservados como referência histórica; as instruções atuais são mantidas em Markdown.

## 1. Preparação e versões

O LynxAutomator prepara ficheiros Excel, descarrega imagens autorizadas do Wildlife
Insights e disponibiliza ferramentas de armadilhagem fotográfica. Não envia
automaticamente o Excel nem as imagens para o Wildbook.

Pode escolher português, espanhol ou inglês na interface. Alguns nomes de separadores,
fases de processamento e mensagens de erro permanecem em inglês. Guarde os resultados
antes de mudar o idioma, pois essa ação reconstrói os formulários.

| Funcionalidade | Completa | Alpha mini |
| --- | --- | --- |
| BIWbE a partir de pasta e catálogo | Sim | Sim |
| Descarga WI e conversão de CSV para Excel | Sim | Sim |
| Folhas de acompanhamento do lince | Sim | Sim |
| Extração de fotogramas | Sim | Sim |
| Módulo de correção de datas dos originais | Sim | Não |
| Módulo de renomeação dos originais | Sim | Não |

A mini também cria ficheiros, executa o gsutil e ajusta as datas dos fotogramas
extraídos. Não é uma versão de apenas leitura e o seu nome não implica assinatura
digital. Nesta branch, as duas interfaces partilham os serviços de processamento;
a mini não apresenta todas as opções da completa. O workflow empacota apenas a
versão completa. Consulte a [avaliação das versões e da interface](architecture.md).

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

## 5. Wildlife Insights → Descarga WI

### Preparar os dados e o acesso

Peça no Wildlife Insights uma exportação com os filtros pretendidos, extraia o
pacote recebido e localize `images.csv`. Siga o guia de utilização e citação que
acompanha essa exportação para obter acesso às imagens. Consulte a
[documentação oficial de descargas privadas](https://www.wildlifeinsights.org/get-started/download/private).

Instale o Google Cloud CLI com `gsutil` disponível no PATH e configure uma conta
autorizada de acordo com o guia. `gsutil version` permite verificar se a ferramenta
está disponível; instalá-la não concede, por si só, acesso ao bucket.

### Descarregar imagens

1. Selecione um CSV com as colunas `location` e `deployment_id`, sem valores vazios.
2. Escolha se pretende separar as imagens em pastas por instalação (`deployment`).
3. Clique em descarregar e escolha o destino.
4. No fim, reveja as quantidades concluídas, ignoradas e falhadas.

As localizações têm de começar por `gs://` e apontar para JPEG (`.jpg` ou `.jpeg`).
A aplicação verifica o conteúdo JPEG, substitui caracteres não suportados nos nomes
e grava com extensão `.JPG`. Não converte outros formatos de imagem para JPEG.

Os ficheiros existentes são ignorados, sem nova validação do conteúdo. As novas
transferências usam pastas temporárias para não confundir descargas incompletas
com imagens terminadas. Se duas localizações diferentes produzirem o mesmo nome
de saída numa execução, o conflito é indicado.

Use **Cancelar**, na barra inferior, para parar depois da transferência em curso.
Espere pelo fim antes de iniciar outra descarga.

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
| Aviso de editor desconhecido ou bloqueio ao abrir | Origem do pacote e política do computador; a mini não substitui a assinatura. Não desative a proteção para experimentar ao acaso. |
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

Abra **Camtrap DP**, carregue `datapackage.json` ou ZIP local (1.x, CSV/CSV.gz) e
marque as espécies usando a pesquisa e as caixas de seleção. Não é necessário
registar um Wildbook. Reveja o número de imagens únicas antes de escolher o destino.
As imagens associadas por evento são opcionais, usam a instalação e o intervalo
temporal e precisam de revisão. O modo local não acede à Internet.

A pasta `camtrap-…` contém as imagens verificadas e `manifest.csv` com IDs, espécies,
nomes e estados. Imagens privadas e vídeos são omitidos. URLs HTTP acessíveis são
transferidos sem credenciais adicionais. Repetir falhadas cria um novo lote.
Cancelar preserva os ficheiros completos e limpa os parciais; uma leitura HTTP em
curso pode esperar até ao seu limite de 20 segundos.

O exemplo sintético local opcional (excluído do Git) tem 366 observações de lince, 247 imagens diretamente
associadas, 300 incluindo eventos e 10 JPEG locais. As fotos não mostram linces.
A validação é estrutural básica, não completa. Esta aba ainda não exporta Excel
Wildbook nem se liga diretamente à API de Agouti.
