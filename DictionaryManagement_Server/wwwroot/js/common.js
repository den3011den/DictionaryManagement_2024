window.ShowToastr = (type, message) => {
    if (type === "success") {
        toastr.success(message, "", { timeOut: 5000 });
    }
    if (type === "error") {
        toastr.error(message, "Ошибка выполнения операции", { timeOut: 5000 });
    }
}

window.ShowSwal = (type, message) => {
    if (type === "success") {
        swal(
            'Успешная операция!',
            message,
            'success'
        )
    }
    if (type === "error") {
        swal(
            'Ошибка операции!',
            message,
            'error'
        )
    }
    if (type === "warning") {
        swal(
            'Предупреждение!',
            message,
            'warning'
        )
    }
    if (type === "loading") {
        swal({
            title: "Операция выполняется ...",
            text: message,           
            type: "info",
            icon: "/images/loading.gif",
            allowOutsideClick: false,
            allowEscapeKey: false,
            allowEnterKey: false,
            showConfirmButton: false,
            showCancelButton: false,
            showCloseButton: false,
            showConfirmButton: false            
        });
        
    }
}

window.CloseSwal = () => {
    swal.close()    
}



//() => {
//    Swal.showLoading()
//},

function ShowDeleteConfirmationModal() {
    $('#deleteConfirmationModal').modal('show');
}

function HideDeleteConfirmationModal() {
    $('#deleteConfirmationModal').modal('hide');
}