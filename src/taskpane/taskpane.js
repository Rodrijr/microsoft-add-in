
var clientID = "f5721a40-33b8-4b2b-8470-44db5b7813fa"// Obtener el objeto Office.js
Office.onReady(function () {
  console.log('Office.js cargado');

  // Agregar evento de inicio de sesión
  Office.context.ui.messageParent({
    action: 'login',
    data: {
      title: 'Iniciar sesión en ServiceNow',
      description: 'Por favor, inicia sesión en ServiceNow para continuar.'
    }
  });

  // Evento de inicio de sesión
  Office.context.ui.addHandler('login', function (event) {
    console.log('Evento login detectado');

    // Verificar si el usuario está autenticado en ServiceNow
    if (Office.context.authenticatedUser.getDisplayName()) {
      console.log('Usuario autenticado: ' + Office.context.authenticatedUser.getDisplayName());

      // Mostrar un mensaje de bienvenida al usuario
      Office.context.ui.messageParent({
        action: 'welcome',
        data: {
          title: 'Bienvenido, ' + Office.context.authenticatedUser.getDisplayName(),
          description: 'Por favor, selecciona una acción para continuar.'
        }
      });

      // Agregar evento de selección de acción
      Office.context.ui.addHandler('actionSelected', function (event) {
        console.log('Evento actionSelected detectado');

        // Obtener la acción seleccionada por el usuario
        var action = event.data.action;

        // Realizar la acción seleccionada (por ejemplo, llamar a una API ServiceNow)
        switch (action) {
          case 'mostrarInformacion':
            mostrarInformacion();
            break;
          default:
            console.log('Acción no reconocida');
        }
      });
    } else {
      console.log('Usuario no autenticado');
    }
  });

  // Función para mostrar información sobre el usuario
  function mostrarInformacion() {
    Office.context.ui.messageParent({
      action: 'mostrarInformacion',
      data: {
        title: 'Información del usuario',
        description: 'Nombre: ' + Office.context.authenticatedUser.getDisplayName()
      }
    });
  }

  // Agregar evento de inicio de sesión con ServiceNow
  document.getElementById('miIframe').addEventListener('load', function () {
    console.log('iFrame cargado');

    // Obtener el objeto del iFrame (ServiceNow)
    var iframe = document.getElementById('miIframe');
    var win = iframe.contentWindow;

    // Agregar evento de inicio de sesión con ServiceNow
    win.addEventListener('message', function (event) {
      console.log('Mensaje recibido desde ServiceNow');

      // Verificar si el mensaje es un evento de inicio de sesión exitoso
      if (event.data.type === 'loginSuccess') {
        console.log('Inicio de sesión en ServiceNow exitoso');

        // Mostrar un mensaje de bienvenida al usuario
        Office.context.ui.messageParent({
          action: 'welcome',
          data: {
            title: 'Bienvenido, ' + event.data.username,
            description: 'Por favor, selecciona una acción para continuar.'
          }
        });

        // Agregar evento de selección de acción
        Office.context.ui.addHandler('actionSelected', function (event) {
          console.log('Evento actionSelected detectado');

          // Obtener la acción seleccionada por el usuario
          var action = event.data.action;

          // Realizar la acción seleccionada (por ejemplo, llamar a una API ServiceNow)
          switch (action) {
            case 'mostrarInformacion':
              mostrarInformacion();
              break;
            default:
              console.log('Acción no reconocida');
          }
        });
      } else if (event.data.type === 'loginError') {
        console.log('Inicio de sesión en ServiceNow fallido');

        // Mostrar un mensaje de error al usuario
        Office.context.ui.messageParent({
          action: 'error',
          data: {
            title: 'Error al iniciar sesión en ServiceNow',
            description: event.data.errorMessage
          }
        });
      }
    });
  });
});