# Importar as funções de cada arquivo
from sorting_in import run_script as run_sorting
from to_be_putaway import run_script as run_putaway
from to_be_picked import run_script as run_picking
from to_be_sorted import run_script as run_sorted
from to_be_packed import run_script as run_packed

# Importar a biblioteca pywin32
import ctypes

# Função para notificar o Agendador de Tarefas que a tarefa foi concluída
def notify_task_completed():
    ctypes.windll.kernel32.SetExitCode(0)

# Executar as funções em sequência
def main():
    functions = [run_sorting, run_putaway, run_picking, run_sorted, run_packed]
    
    for i, function in enumerate(functions, start=1):
        result = function()
        print(f"Função {i}: {result}")
        
        if not result:
            print("Alguma função retornou False. Encerrando o programa.")
            notify_task_completed()  # Notificar o Agendador de Tarefas
            return

if __name__ == "__main__":
    main()
