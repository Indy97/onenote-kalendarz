// Konfiguracja wtyczki. Uzupełnij clientId po rejestracji aplikacji w Entra ID
// (instrukcja: README.md, krok 2).
window.KALENDARZ_CONFIG = {
  clientId: "WSTAW-CLIENT-ID-APLIKACJI",
  tenantId: "2eaef7c3-6a44-4ff4-a487-cb440a8749e1",
  // Nazwa listy w Microsoft To Do, w której trzymane są zadania.
  todoListName: "Kalendarz OneNote",
  // Strefa czasowa w formacie Windows (Graph wymaga tej nazwy).
  timeZone: "Central European Standard Time",
  defaultReminderTime: "08:00",
  scopes: ["User.Read", "Tasks.ReadWrite"]
};
