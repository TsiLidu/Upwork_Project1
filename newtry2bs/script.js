/* bar chart */
        
            var chBar = document.getElementById("barChart");
            if (chBar) {
              new Chart(chBar, {
              type: 'bar',
              data: {
                labels: ["S", "M", "T", "W", "T", "F", "S"],
                datasets: [{
                  data: [589, 445, 483, 503, 689, 692, 634],
                  backgroundColor: colors[0]
                },
                {
                  data: [639, 465, 493, 478, 589, 632, 674],
                  backgroundColor: colors[1]
                }]
              },
              options: {
                legend: {
                  display: false
                },
                scales: {
                  xAxes: [{
                    barPercentage: 0.4,
                    categoryPercentage: 0.5
                  }]
                }
              }
              });
            } 
          