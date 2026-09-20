# Implantação – PE | Planejamento Empresarial

## Estrutura já preparada

- Turma: `1º MTEC - Administração - Mairinque`
- Componente: `PE`
- Banco Central: `1QpyTzbc2P0F-LrGw7fIe7m0MOE10rbVSFAyCiWrgCNI`
- Pasta raiz: `1oUeynmQmDBTemiDZCXn-keIIw6Ga3TAA`
- Materiais: `1lg2ToRMQ-XvvBCrBQehbZZdAqFIUbYpy`
- Entregas: `1tGnmfyLALLVaTTZwzoDc_oCCSJnAQZlZ`

## 1. Backend público (alunos)

1. Abra https://script.google.com e crie um novo projeto chamado `PE - Backend Alunos`.
2. Apague o conteúdo padrão de `Code.gs`.
3. Copie integralmente o arquivo `backend/Code.gs` deste repositório para `Code.gs`.
4. Salve.
5. Vá em **Implantar > Nova implantação**.
6. Tipo: **Aplicativo da Web**.
7. Executar como: **Eu**.
8. Quem tem acesso: **Qualquer pessoa**.
9. Implante e autorize as permissões solicitadas.
10. Copie a URL terminada em `/exec`.

## 2. Painel administrativo (professor)

1. Abra o projeto Apps Script `PE - Administrativo` já criado no Drive, ou crie um projeto novo com esse nome.
2. Em `Code.gs`, copie integralmente `backend-admin/Code.gs` deste repositório.
3. Crie/abra o arquivo HTML `Admin` e copie integralmente `backend-admin/Admin.html`.
4. Salve.
5. Vá em **Implantar > Nova implantação**.
6. Tipo: **Aplicativo da Web**.
7. Executar como: **Eu**.
8. Quem tem acesso: preferencialmente **Somente eu** ou a opção restrita disponível na conta institucional/pessoal.
9. Implante e copie a URL terminada em `/exec`.

## 3. Finalização

Depois das duas implantações, preencher:

- `CONFIG!B7` com a URL do backend público (`WEB_APP_URL`).
- `CONFIG!B8` com a URL do painel administrativo (`WEB_APP_ADMIN_URL`).
- `config.js` com a URL do backend público em `apiUrl`.

Após isso, testar:

1. página principal carrega sem erro;
2. atividade publicada aparece no site;
3. envio de teste gera protocolo;
4. arquivo enviado aparece na pasta `ENTREGAS`;
5. registro aparece na aba `ENTREGAS`;
6. painel do professor lista a entrega.
