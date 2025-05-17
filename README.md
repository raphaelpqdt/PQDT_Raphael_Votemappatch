# Predadores Votemap Patch

Uma aplicação em Python com interface gráfica (Tkinter + ttkbootstrap) desenvolvida para monitorar logs em tempo real do servidor de jogo *Arma Reforger*, atualizando automaticamente os arquivos JSON de configuração conforme o mapa vencedor da votação e reiniciando o serviço do servidor.

## ✨ Funcionalidades

- Monitoramento em tempo real de arquivos `console.log`
- Seleção automática do mapa vencedor da votação (`Winner: [n]`)
- Atualização do campo `scenarioId` no JSON do servidor
- Reinício automático do serviço do servidor (opcional)
- Reversão automática do JSON para voltar ao votemap ao final da partida
- Exibição ao vivo dos conteúdos JSON
- Suporte a filtro de logs e troca de tema

## 🚀 Como usar

1. Execute o arquivo `.exe` (ou `.py`) da aplicação.
2. Selecione:
   - A pasta dos logs
   - O JSON do servidor
   - O JSON do votemap
   - (Opcional) o serviço do Windows correspondente ao servidor
3. Acompanhe a votação e deixe o sistema automatizar o processo, sem problemas de JIP_ERROR_8

## 🛠️ Tecnologias

- Python 3.13
- `tkinter` + `ttkbootstrap`
- `pystray`
- `win32com.client` (para gerenciamento de serviços)
- `subprocess`, `json`, `os`, `threading`

## 📄 Licença

Distribuído sob a **Creative Commons Zero v1.0 Universal (CC0)**.  
**Proibido uso comercial** e **modificações devem manter o projeto open source**.

---

Desenvolvido com dedicação por **@raphaelpqdt** 🎮💻 
https://www.linkedin.com/in/raphaelpqdt/
https://www.instagram.com/pqdtraphael/

