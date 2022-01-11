function digital_root(n) {
  var num = n.toString().split("");
  var resultado = 0;

        for(x=0; x <= num.length-1; x++ ){
          resultado += parseInt(num[x]);
        }
       if(resultado.toString().split("").length > 1){
            resultado = digital_root(resultado);
        }

        
  return resultado;
}