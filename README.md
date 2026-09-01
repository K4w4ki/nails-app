# Nails by Nascimento - Sistema de Agendamento

Sistema de agendamento para manicure desenvolvido como projeto piloto acadêmico.

## 📋 Visão Geral

- **Cliente**: agenda serviços, consulta horários, cancela e reagenda
- **Profissional (Carol)**: gerencia agendamentos, disponibilidade, serviços e faturamento
- **Tecnologia**: Firebase (Authentication + Firestore)

## 🚀 Como Configurar

### 1. Criar Projeto Firebase

1. Acesse [Firebase Console](https://console.firebase.google.com/)
2. Clique em "Criar projeto"
3. Dê um nome (ex: "nails-by-nascimento")
4. Siga as instruções até a criação

### 2. Ativar Serviços

**Authentication:**
- No menu esquerdo, clique em "Authentication"
- Clique em "Começar"
- Ative "E-mail/senha"

**Firestore:**
- No menu esquerdo, clique em "Firestore Database"
- Clique em "Criar banco de dados"
- Escolha "Modo de teste" (inicialmente)
- Selecione a região mais próxima

### 3. Criar Administrador

1. No Authentication, vá para "Usuários"
2. Clique em "Adicionar usuário"
3. Preencha:
   - E-mail: carol@exemplo.com
   - Senha: (crie uma senha segura)
4. Copie o UID do usuário criado

### 4. Criar Coleção de Administradores

No Firestore, crie a coleção:


### 5. Configurar Credenciais

1. No Firebase Console, vá em "Configurações do projeto" → "Geral"
2. Role até "Seus apps" e copie o objeto de configuração
3. No arquivo `firebase-config.js`, substitua os placeholders

### 6. Aplicar Security Rules

1. No Firestore, vá em "Regras"
2. Cole o conteúdo de `firestore.rules`
3. Clique em "Publicar"

### 7. Publicar o Site

Opções simples:
- **GitHub Pages**: crie um repositório e ative o Pages
- **Netlify/Vercel**: arraste a pasta do projeto
- **Local**: use extensão Live Server no VS Code

## 📁 Estrutura do Projeto


## 🔄 Fluxos Principais

### Agendamento
1. Cliente escolhe serviço
2. Seleciona data e horário
3. Preenche dados e envia
4. Status: **pendente**

### Administração
1. Carol faz login
2. Visualiza solicitações
3. Aprova/recusa
4. Gerencia disponibilidade
5. Edita serviços

### Cliente
1. Consulta pelo WhatsApp
2. Visualiza status
3. Cancela ou reagenda
4. Limite de 3 reagendamentos

## 🔐 Status dos Agendamentos

| Status | Descrição |
|--------|-----------|
| pendente | Aguardando análise da Carol |
| aguardando-sinal | Aprovado, aguardando depósito |
| confirmado | Agendamento confirmado |
| concluido | Atendimento realizado |
| recusado | Não aprovado pela Carol |
| cancelado | Cancelado pelo cliente |

## 📱 WhatsApp

Os links utilizam `wa.me` para:
- Cobrança de sinal
- Lembretes
- Comunicação direta

## 🛠️ Tecnologias

- HTML/CSS (design original preservado)
- JavaScript (vanilla)
- Firebase Authentication
- Cloud Firestore

## 📝 Notas

- Este é um **projeto piloto acadêmico**
- Os dados são persistidos no Firestore
- Não armazene senhas no código ou no Firestore
- O administrador é definido pela coleção `admins`

## 📄 Licença

Uso acadêmico e demonstrativo.