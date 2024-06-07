import tkinter as tk
import can
from can.interface import Bus
from time import sleep as wait
import can.interfaces.vector
import Diag_testing

def focus_next(event):
    event.widget.tk_focusNext().focus()
    return "break"

def save_text():
    save_text = [text_box1.get("1.0", tk.END).strip() + text_box2.get("1.0", tk.END).strip() + text_box3.get("1.0", tk.END).strip() + text_box4.get("1.0", tk.END).strip()]
    text_box1.delete("1.0", tk.END)
    text_box2.delete("1.0", tk.END)
    text_box3.delete("1.0", tk.END)
    text_box4.delete("1.0", tk.END)
    label.config(text="send Data:{}".format(save_text))

def button_click1():
    can_bus = can.interface.Bus(bustype='vector', app_name='CANalyzer', channel=0, bitrate=500000)
    service_id = 0x7C6
    data = [0x3, 0x22, 0x00, 0x05, 0x0, 0x0, 0x0, 0x0]
    send_msg = can.Message(arbitration_id=service_id, data=data, is_extended_id=False)
    can_bus.send(send_msg)
    count = 0
    while True:
        recv_msg =can_bus.recv()
        if recv_msg is not None and recv_msg.arbitration_id == 0x7CE:
            label.config(text="Cluster SW Version:{}".format(recv_msg.data))
            break
        else:
            if count < 100:
                count = count + 1
            else:
                print("No response message received")
                break
    can_bus.shutdown()
    
def button_click2():
    can_bus = can.interface.Bus(bustype='vector', app_name='CANalyzer', channel=0, bitrate=500000)
    service_id = 0x7C6
    data = [0x3, 0x22, 0xf1, 0x87, 0x0, 0x0, 0x0, 0x0]
    send_msg = can.Message(arbitration_id=service_id, data=data, is_extended_id=False)
    can_bus.send(send_msg)
    count = 0
    while True:
        recv_msg =can_bus.recv()
        if recv_msg is not None and recv_msg.arbitration_id == 0x7CE:
            label.config(text="Product No.:{}".format(recv_msg.data))
            break
        else:
            if count < 100:
                count = count + 1
            else:
                print("No response message received")
                break
    can_bus.shutdown()

def button_click3():
    can_bus = can.interface.Bus(bustype='vector', app_name='CANalyzer', channel=0, bitrate=500000)
    service_id = 0x7C6
    data = [0x3, 0x22, 0x00, 0x01, 0x0, 0x0, 0x0, 0x0]
    send_msg = can.Message(arbitration_id=service_id, data=data, is_extended_id=False)
    can_bus.send(send_msg)
    count = 0
    while True:
        recv_msg =can_bus.recv()
        if recv_msg is not None and recv_msg.arbitration_id == 0x7CE:
            label.config(text="Product No.:{}".format(recv_msg.data))
            break
        else:
            if count < 100:
                count = count + 1
            else:
                print("No response message received")
                break
    can_bus.shutdown()


if __name__ == '__main__':    

    window = tk.Tk()
    window.geometry("450x300")
    window.configure(bg="white")
    window.title("Visteon.Diagnostic for cluster")

    button1 = tk.Button(window, text="Version", command=button_click1, width = 9, height = 2,font=("Helvetica", 12))
    button1.pack()
    
    button2 = tk.Button(window, text="product No.", command=button_click2, width = 9, height = 2,font=("Helvetica", 12))
    button2.pack(pady=5)  

    label = tk.Label(window, text="Value")
    label.pack(padx=3, pady=20) 

    frame1 = tk.Frame(window)
    frame1.pack()

    text_box1 = tk.Text(frame1, width=9, height=1)
    text_box1.pack(side=tk.LEFT)
    
    label1 = tk.Label(frame1, text="data1")
    label1.pack(side=tk.LEFT)
    
    text_box1.bind("<Tab>", focus_next)

    text_box2 = tk.Text(frame1, width=9, height=1)
    text_box2.pack(side=tk.LEFT)
    
    label2 = tk.Label(frame1, text="data2")
    label2.pack(side=tk.LEFT)
    
    text_box2.bind("<Tab>", focus_next)
    
    frame2 = tk.Frame(window)
    frame2.pack()

    text_box3 = tk.Text(frame2, width=9, height=1)
    text_box3.pack(side=tk.LEFT)
    
    label3 = tk.Label(frame2, text="data3")
    label3.pack(side=tk.LEFT)
    
    text_box3.bind("<Tab>", focus_next)
    
    text_box4 = tk.Text(frame2, width=9, height=1)
    text_box4.pack(side=tk.LEFT)
    
    label4 = tk.Label(frame2, text="data3")
    label4.pack(side=tk.LEFT)
    
    text_box4.bind("<Tab>", focus_next)
    
    save_button = tk.Button(window, text="send value", command=save_text)
    save_button.pack(padx=3, pady=20)
    
    can_bus = can.interface.Bus(bustype='vector', app_name='CANalyzer', channel=0, bitrate=500000)

    Diag_testing.enter_input_session(can_bus,"Extended")
    Diag_testing.ASK_security(can_bus)

    window.mainloop()
    
