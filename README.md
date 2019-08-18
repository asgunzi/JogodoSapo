# JogodoSapo

Implementação em Excel – VBA do “Jogo do Sapo”.


Faça os três sapos da esquerda trocarem de posição com os da esquerda.

![](https://ferramentasexcelvba.files.wordpress.com/2019/08/sapo1.png?w=656)

Clique no sapo para movimentá-lo. Ex. Cliquei no terceiro sapo, e ele avança uma casa se esta estiver livre, e duas se a segunda estiver livre (pulando quem estiver na frente).

![](https://ferramentasexcelvba.files.wordpress.com/2019/08/sapo2.png?w=656)
É necessário ativar macros para rodar o jogo.

Em termos de VBA, o truque é mais ou menos simples.

Em linhas gerais, cada sapo é uma imagem, com um nome diferente.

Podemos selecionar o sapo desejado adaptando o comando abaixo:

    ActiveSheet.Shapes.Range(Array(“Sapo1”)).Select

Uma vez selecionado, podemos posicioná-lo com as propriedades Top e Left (equivalente ao eixo y e x).

    Selection.Top = y0

    Selection.Left = x0 + (i – 1) * delta


Objetivo: Trocar os sapos de posição.
![](https://ferramentasexcelvba.files.wordpress.com/2019/08/sapo3.jpg?w=656)


Ideias técnicas com uma pitada de filosofia: https://ideiasesquecidas.com

Ferramentas Excel-VBA: https://ferramentasexcelvba.wordpress.com/
