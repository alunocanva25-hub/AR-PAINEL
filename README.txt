DSYSTEM AR PAINEL V1.0.0.9

Atualizações:
- Ícone do menu removido.
- Quando o menu está recolhido, aparece apenas DS.
- Menu lateral abre ao passar o mouse e fecha ao sair.
- Mantidas as funções da V1.0.0.7.

Execute RODAR_LOCAL.bat ou INICIAR_REDE.bat.


V1.0.1.3.3_RECEBER_API_SYNC_DELETE
- Atualizado com base na V1.0.1.2 enviada pelo usuário.
- Adicionado botão Receber API.
- Botão Atualizar recebe novos registros da API já configurada.
- Ao excluir no painel, se o registro veio da API, tenta excluir também na API.


V1.0.1.4_OPERADOR
- Painel recebe operador_usuario, operador_nome e operador_perfil da API.
- Adicionada coluna Enviado por.
- Adicionado filtro por operador.


V1.0.1.5_COLUNAS_DATA_MARCAR_TODOS
- Corrigido alinhamento das colunas: Nome Cliente, Enviado por, Nome do Arquivo e Status.
- Adicionado checkbox Marcar todos.
- Campos Data inicial/Data final adicionados aos filtros.
- Mantido botão Baixar Selecionados existente.


V1.0.1.6_LOGIN_CONFIG
- Tela de login adicionada.
- Usuário padrão:
  admin / admin123
- Menu Configuração.
- Admin cria usuários.
- Admin troca senha e exclui usuários.

V1.0.1.7_LOGIN_API_CONFIG
- Login do painel usa API central: https://dsystem-ar-api.onrender.com
- Adicionado botão Logout/Sair.
- Configuração cria/exclui usuários na API.
- Admin troca senha de usuários pela API.
- X do overlay de configuração centralizado.
- Área Criar usuário com recolher/exibir.


V1.0.1.8_LOGIN_API_FUNCIONAL
- Corrigido Failed to fetch no login quando havia URL antiga/salva errada.
- Painel força URL padrão da API Render: https://dsystem-ar-api.onrender.com
- Se a API não responder, admin/admin123 entra em modo local temporário.
- Corrige URL local https://IP para http://IP ao receber API.
- Mantém Logout, X centralizado e recolher Criar usuário.

V1.0.1.9_AUTO_REFRESH
- Adicionada opção em Configuração: Atualização automática.
- Permite ativar/desativar atualização automática.
- Intervalos disponíveis: 5, 10 ou 15 minutos.
- Quando ativo, o painel consulta a API automaticamente para verificar novos uploads.
- Bloco de atualização automática também possui Recolher/Exibir.

V1.0.2.0_USER_INITIALS
- Avatar superior mostra as duas iniciais do nome do usuário logado.
- Se não houver nome, usa as iniciais do login.

V1.0.2.1_USERS_COLLAPSE
- Adicionada opção Recolher/Exibir na seção Usuários dentro de Configuração.
- A lista de usuários agora fica em área com rolagem para não crescer demais.

V1.0.2.2_BASE_XLSX_API_CACHE
- Painel não processa mais XLSX localmente.
- Painel envia XLSX para a API.
- API processa e salva no PostgreSQL.
- Evita estourar memória do Render no painel.
- Mantém as telas e botões existentes.

V1.0.2.3_BASE_FIND_FLEX
- Painel compatível com API V1.0.0.9.
- Após processar XLSX, solicita reindexação da base na API.
- Busca na base passa a aceitar diferenças de pontuação, hífen, espaços e formatação.
