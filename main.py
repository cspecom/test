# V1.0 features:
# Tự động bật máy theo danh sách nhập bằng file text, bật theo thứ tự từ trên xuống dưới hoặc random thứ tự các máy.
# Tính năng thiết lập giới hạn số lượng máy ảo chay cùng lúc
# Nếu các máy được shutdown bằng phần mềm bên trong thành công -> tiếp tục khởi phiên máy ảo mới sau thời gian chờ
# Tính năng Hẹn giờ chạy tối đa của máy để force-shutdown, các máy được thiết lập  bộ đếm giờ riêng
# Chương trình dừng lại khi các máy đã được bật hết 1 lượt và không còn máy nào đang chạy

import subprocess
import time
import random

# Step 1: Read virtual machine information from a text file in the format "group/VMname"
def read_vm_info(file_path):
    vm_info = []
    with open(file_path, "r") as file:
        lines = file.readlines()
        for line in lines:
            group, vm_name = line.strip().split("/")
            vm_info.append((group, vm_name))
    random.shuffle(vm_info) # Xáo trộn thứ tự các phần tử trong list
    return vm_info

# Step 2: Function to check the number of running virtual machines
def get_running_vm_count():
    vm_list = subprocess.check_output(["VBoxManage", "list", "runningvms"]).decode().splitlines()
    return len(vm_list)

# Step 3: Start new virtual machines if the running count is less than x
def start_new_vms(vm_info, max_concurrent_vms, vm_status_dict, start_time_dict):
    count = get_running_vm_count()
    print(f"The number of VM currently running: {count}")
    if count < max_concurrent_vms:
        for group, vm_name in vm_info:
            if count >= max_concurrent_vms:
                break
            if vm_status_dict[vm_name]:
                # The VM has already been started, skip it
                continue
            vm_info_output = subprocess.check_output(["VBoxManage", "showvminfo", vm_name]).decode()
            if "running" not in vm_info_output:
                subprocess.Popen(["VBoxManage", "startvm", vm_name])
                count += 1
                time.sleep(45)
                # subprocess.run(["VBoxManage", "controlvm", vm_name, "keyboardputscancode","3b","02"])
                # print(f"Successfully disabled Mouse Integration for VM {vm_name}")
                vm_status_dict[vm_name] = True # Mark the VM as started in the dictionary
                start_time_dict[vm_name] = time.time() # Record the start time of the VM
                time.sleep(15) # Wait 60 seconds before starting the next VM

def check_running_time_dict(vm_info, time_to_run_dict, start_time_dict):
    running_vm_count = 0 # Track the number of VMs ever turned on
    running_vm_info = [] # Track the info of running VMs
    for group, vm_name in vm_info:
        vm_info_output = subprocess.check_output(["VBoxManage", "showvminfo", vm_name]).decode()
        if "running" in vm_info_output:
            running_vm_info.append((group, vm_name))
            start_time = start_time_dict.get(vm_name, None)
            if start_time is not None:
                running_vm_count += 1

    while True:
        time.sleep(20) # Wait for 20 seconds before rechecking
        for group, vm_name in running_vm_info:
            vm_info_output = subprocess.check_output(["VBoxManage", "showvminfo", vm_name]).decode()
            if "running" in vm_info_output:
                start_time = start_time_dict.get(vm_name, None)
                if start_time is not None:
                    elapsed_time = time.time() - start_time
                    if elapsed_time >= time_to_run_dict[vm_name]:
                        subprocess.Popen(["VBoxManage", "controlvm", vm_name, "shutdown"])
                        print(f"VM '{vm_name}' has been successfully powered off.")
                        running_vm_count -= 1 # Decrease the count of running VMs
                    else:
                        print(f"VM '{vm_name}' is still running after {elapsed_time} seconds. Waiting for 20 seconds before rechecking.")
            else:
                print(f"VM '{vm_name}' has been turned off")
                running_vm_count = get_running_vm_count() # Count of running VMs

        if running_vm_count < max_concurrent_vms:
            break
        # else:
        #     running_vm_info = [(group, vm_name) for group, vm_name in running_vm_info if subprocess.check_output(["VBoxManage", "showvminfo", vm_name]).decode().strip().endswith("running")]

# Main program
if __name__ == "__main__":
    file_path = "list.txt"
    start_time_dict = {}

    # Set Maximum number of concurrent running VMs
    max_concurrent_vms = 2

    vm_info = read_vm_info(file_path)
    vm_names = [vm[1] for vm in vm_info]

    # Set max time to run for all VMs (s)
    time_to_run_dict = {vm_name: 100 for vm_name in vm_names}

    # Initialize the VM status dictionary
    vm_status_dict = {vm_name: False for _, vm_name in vm_info}

    while True:
        start_new_vms(vm_info, max_concurrent_vms, vm_status_dict, start_time_dict)
        check_running_time_dict(vm_info, time_to_run_dict, start_time_dict)
        print("Prepare to initialize the next VM if available")
        if all(vm_status_dict.values()) and all(
                "powered off" in subprocess.check_output(["VBoxManage", "showvminfo", vm_name]).decode() for _, vm_name
                in vm_info):
            print("All VMs have been started and checked for the specified time and then turned off.")
            break

