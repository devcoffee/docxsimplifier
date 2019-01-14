![](https://devcoffee.com.br/wp-content/uploads/2017/06/logo_atualizado.gif)

O DocxSimplifier tem como objetivo prover o tratamento e simplificação de arquivos .docx visando facilitar o consumo dos mesmos por APIs e aplicações, como as do próprio BrERP.

# O formato .docx

O formato .docx é, em verdade, um pacote com diversos arquivos .xml, que definem estilos entre os conteúdos, bem como a posição da ultima edição, e condições de checagem ortográfica, funções que em si, não tem muita utilidade para as APIs de consumo, ao contrário, podem comprometer o consumo dos documentos, uma vez que as tags do XML podem quebrar váriaveis de contexto.

Por exemplo, o .xml de um arquivo .docx simples, é formado de maneira similar a essa:

```xml
    <w:body>
        <w:p>
            <w:r>
                <w:t xml:space="preserve">Use BrERP! (: </w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838" />
            <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0" />
            <w:cols w:space="708" />
            <w:docGrid w:linePitch="360" />
        </w:sectPr>
    </w:body>
```

No entanto, conforme o usuário adiciona Estilos no Texto, o mesmo é fracionado entre as tags no .xml, o que pode comprometer o consumo do documento por APIs. Além disso, o próprio Microsoft Word (Libreoffice Write, WPS Writter, GDocs, ou qualquer outro) também adiciona tags de Profiler de Erro e ultima posição do Cursor, ambas não desejadas.

O mesmo documento, quando formatado, demonstra o seguinte .xml:

```xml
    <w:body>
        <w:p w:rsidR="00CA1A01" w:rsidRDefault="00CA1A01" w:rsidP="00CA1A01">
            <w:r w:rsidRPr="00CA1A01">
                <w:rPr>
                    <w:color w:val="FF0000"/>
                </w:rPr>
                <w:t>Use</w:t>
            </w:r>
            <w:r>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:proofErr w:type="spellStart"/>
            <w:r>
                <w:t>B</w:t>
            </w:r>
            <w:r w:rsidRPr="00CA1A01">
                <w:rPr>
                    <w:highlight w:val="yellow"/>
                </w:rPr>
                <w:t>rER</w:t>
            </w:r>
            <w:r>
                <w:t>P</w:t>
            </w:r>
            <w:proofErr w:type="spellEnd"/>
            <w:r>
                <w:t xml:space="preserve">! (: </w:t>
            </w:r>
        </w:p>
        <w:p w:rsidR="00440F53" w:rsidRPr="00CA1A01" w:rsidRDefault="00440F53" w:rsidP="00CA1A01">
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
        </w:p>
        <w:sectPr w:rsidR="00440F53" w:rsidRPr="00CA1A01">
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
            <w:cols w:space="708"/>
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>
    </w:body>

```

Com as tags de texto <w:t> divididas entre as tags de estilo, correção, e outras, podem surgir complicações, ou mesmo um aumento desnecessário da complexidade do algoritmo de consumo de sua API.

# Utilização

Com o intuito de facilitar o consumo dos documentos .docx por APIs,  o DocxSimplifier é um aplicativo de console, que deve ser chamado a partir do CMD, Powershell, Bat File, ou qualquer meio que tenha interação com o Console do Windows.

## Argumentos de Linha de Comando

Existem dois argumentos que podem ser chamados para executar o DocxSimplifier:

- docLocation: A localização do arquivo a ser simplificado. Caso ele esteja na pasta da aplicação, pode ser passado apenas o nome do arquivo.
- --removeStyles: Essa flag irá indicar para o DocxSimplifier que o arquivo deve ser reescrito, todo com um estilo padrão. Atenção: Se essa flag for passada, todos os paragrafos do documento serão reescritos em Arial 12.

## Exemplo prático: O Template .docx do BrERP

Dentro do BrERP é possível inserir um template de documento .docx, contendo variáveis que serão substituidas por informações do sistema. Esta é uma ferramenta poderosa e pratica, que permite aos clientes a geração de documentos personalisados de maneira prática.

### Documento formatado corretamente

Um exemplo de documento formatado corretamente pode ser visto na imagem abaixo:

![](DocxSimplifier/documents/documentoFormatado.png)

A API de consumo de .docx do BrERP substituirá as variáveis de contexto (texto circundado entre @) pelos respectivos valores desejados. No entanto, a API de consumo trabalha analisando o conteúdo das tags <w:t>. E, ao analisarmos o arquivo .xml dentro do .docx:

```xml
<w:body>
    <w:p w:rsidR="000904EC" w:rsidRPr="000904EC" w:rsidRDefault="000904EC" w:rsidP="000904EC">
        <w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
        </w:pPr>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>Empresa:@</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>AD_Client_ID</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>&lt;</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>Name</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>&gt;@</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>Org</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>:@</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>AD_Org_ID</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>&lt;</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>Name</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>&gt;@Orçamento: @</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>DocumentNo@Emissão</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>:@</w:t>
        </w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>DateOrdered</w:t>
        </w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r w:rsidRPr="000904EC">
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                <w:color w:val="FF0000"/>
                <w:sz w:val="24"/>
                <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>@</w:t>
        </w:r>
    </w:p>
    <w:p w:rsidR="00FC0847" w:rsidRPr="000904EC" w:rsidRDefault="00FC0847" w:rsidP="000904EC"/>
    <w:sectPr w:rsidR="00FC0847" w:rsidRPr="000904EC">
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
        <w:cols w:space="708"/>
        <w:docGrid w:linePitch="360"/>
    </w:sectPr>
</w:body>
```

As Tags de Error Profiler e de RSID, adicionadas pelo Microsoft Word, fracionam as tags de texto <w:t> em várias partes, dificultando o trabalho de nossa API. 
Para corrigir isso, basta executar o DocxSimplifier:

```batch
DocxSimplifier.exe DocumentoTeste.docx
```

Agora, o arquivo .xml tem a seguinte aparencia:

```xml
    <w:body>
        <w:p>
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>Empresa:@AD_Client_ID&lt;Name&gt;@Org:@AD_Org_ID&lt;Name&gt;@Orçamento: @DocumentNo@Emissão:@DateOrdered@</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838" />
            <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0" />
            <w:cols w:space="708" />
            <w:docGrid w:linePitch="360" />
        </w:sectPr>
    </w:body>
```

Agora, a tag <w:t> não estão mais fracionadas entre tags de Error Profile e RSID, possibilitando o trabalho correto da API de consumo.

### Documento formatado incorretamente

Existem casos que apenas retirar as tags de Error Profile e RSID não é suficiente para manter  a integridade das tags <w:t>. Um exemplo muito comum, é quando o estilo do texto é diferente no começo e no fim. Quando isso acontece com o texto de uma variável de ambiente, a API fica incapaz de atuar sobre ela. Isso acontece na imagem abaixo:

![Documento Mal Formatado](DocxSimplifier/documents/documentoMalFormatado.png)

Neste exemplo, o texto dentro das variáveis de contexto aparece com cores de destaque. Isso gera grande fragmentação nas tags <w:t>, afinal, elas são dividias dentro das tags de estilo.
Mesmo no arquivo .xml que já passou pelo processo de remoção de tags de Error Profile e RSID, pode se notar o fracionamento:

```xml
    <w:body>
        <w:p>
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>Empresa:@</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                    <w:highlight w:val="yellow" />
                </w:rPr>
                <w:t>AD_Client_ID</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>&lt;Name&gt;@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>Org:@AD_Org_ID&lt;</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                    <w:highlight w:val="yellow" />
                </w:rPr>
                <w:t>Name</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>&gt;@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>Orçamento: @</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                    <w:highlight w:val="yellow" />
                </w:rPr>
                <w:t>DocumentNo</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>Emissão:@</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                    <w:highlight w:val="yellow" />
                </w:rPr>
                <w:t>DateOrdered</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
                    <w:color w:val="FF0000" />
                    <w:sz w:val="24" />
                    <w:szCs w:val="24" />
                </w:rPr>
                <w:t>@</w:t>
            </w:r>
        </w:p>
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838" />
            <w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0" />
            <w:cols w:space="708" />
            <w:docGrid w:linePitch="360" />
        </w:sectPr>
    </w:body>
```

Em situações como essa, recomenda-se utilizar a flag --removeStyles para reescrever o documento, mantendo um estilo padrão. Assim, a API de consumo pode processá-lo normalmente. Para fazer, isso:

```batch
DocxSimplifier.exe DocumentoTeste.docx --removeStyles
```

Como Resultado, obtém-se o documento reescrito em um único padrão, e o .xml já simplificado:

![DocumentoReescrito](/DocxSimplifier/documents/documentoReescrito.png)


```xml
    <w:body>
        <w:p>
            <w:pPr>
                <w:jc w:val="left"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                    <w:b w:val="0"/>
                    <w:sz w:val="24"/>
                </w:rPr>
                <w:t xml:space="preserve">Empresa:@AD_Client_ID&lt;Name&gt;@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="left"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                    <w:b w:val="0"/>
                    <w:sz w:val="24"/>
                </w:rPr>
                <w:t xml:space="preserve">Org:@AD_Org_ID&lt;Name&gt;@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="left"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                    <w:b w:val="0"/>
                    <w:sz w:val="24"/>
                </w:rPr>
                <w:t xml:space="preserve">Orçamento: @DocumentNo@</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="left"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                    <w:b w:val="0"/>
                    <w:sz w:val="24"/>
                </w:rPr>
                <w:t xml:space="preserve">Emissão:@DateOrdered@</w:t>
            </w:r>
        </w:p>
    </w:body>
```

# Uso e Licença
Este programa é distribuído sobre a licença *GNU GPL 3.0* na **expectativa de ser útil**, mas  **sem qualquer garantia**; sem mesmo a garantia implícita de **comercialização** ou de **adequação a qualquer propósito particular**.
Consulte a Licença Pública Geral GNU para obter mais detalhes.
Sinta-se livre para indicar erros e apontar soluções.